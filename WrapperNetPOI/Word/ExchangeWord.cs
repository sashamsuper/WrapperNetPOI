#define DEBUG
/* ==================================================================
Copyright 2020-2023 sashamsuper
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
==========================================================================*/
using NPOI.HWPF;
using NPOI.POIFS.FileSystem;
using NPOI.XWPF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Analysis;

namespace WrapperNetPOI.Word
{
    public interface IExchangeWord : IExchange { }

    public abstract class WordExchange<Tout> : IExchangeWord
    {
        protected WordExchange(ExchangeOperation exchange, IProgress<int> progress = null)
        {
            ExchangeOperationEnum = exchange;
            ProgressValue = progress;
        }

        //public List<List<string[]>> Tables { set; get; } = new List<List<string[]>>();
        public IProgress<int> ProgressValue { get; set; }
        public ILogger Logger { get; set; }
        public ExchangeOperation ExchangeOperationEnum { get; set; }
        public Action ExchangeValueFunc { get; set; }
        public List<Tout> ExchangeValue { set; get; }
        public bool CloseStream { get; set; } = true;
        public WordDoc Document { set; get; }
        public string Password { set; get; }

        public void DeleteValue()
        {
            throw new NotImplementedException();
        }

        public void GetInternallyObject(Stream tmpStream, bool addNew)
        {
            if (addNew)
            {

            }
            else
            {
                object doc = null; // ќбъект дл€ хранени€ HWPFDocument или XWPFDocument


                // --- ѕопытка открыть как HWPF (.doc) ---
                // ћы объ€вл€ем MemoryStream здесь, чтобы иметь возможность утилизировать его в finally,
                // если HWPFDocument не возьмет его во владение.
                MemoryStream memoryStreamForHwpf = null;
                POIFSFileSystem nfs = null; // POIFSFileSystem не IDisposable, поэтому без "using"
                try
                {
                    tmpStream.Position = 0; // ”бедитьс€, что оригинальный поток находитс€ в начале дл€ копировани€
                    memoryStreamForHwpf = new MemoryStream();
                    tmpStream.CopyTo(memoryStreamForHwpf);
                    memoryStreamForHwpf.Position = 0; // —бросить скопированный поток дл€ чтени€ POIFSFileSystem

                    nfs = new POIFSFileSystem(memoryStreamForHwpf); // POIFSFileSystem берет на себ€ владение memoryStreamForHwpf

                    // ѕровер€ем наличие специфичной дл€ HWPF записи "WordDocument"
                    if (nfs.Root.HasEntry("WordDocument"))
                    {
                        doc = new HWPFDocument(nfs); // HWPFDocument (IDisposable) берет на себ€ владение nfs.
                                                     // ѕри утилизации HWPFDocument, он утилизирует nfs,
                                                     // который в свою очередь утилизирует memoryStreamForHwpf.
                                                     // ¬ажно: ќбнул€ем memoryStreamForHwpf и nfs, чтобы блок finally не пыталс€ их утилизировать,
                                                     // так как их владение было передано HWPFDocument.
                        memoryStreamForHwpf = null;
                        nfs = null;
                    }
                    else
                    {
                        // Ёто действительный файл POIFS, но не документ Word.
                        // memoryStreamForHwpf все еще принадлежит nfs.
                        // nfs сам не IDisposable, но memoryStreamForHwpf будет утилизирован в блоке finally.
                        Logger?.Debug("POIFSFileSystem успешно создан, но не содержит 'WordDocument' (попробуем XWPF)");
                    }
                }
                catch (Exception e)
                {
                    // ѕерехватывает ошибки копировани€ MemoryStream или создани€ POIFSFileSystem.
                    Logger?.Error("ќшибка при попытке открыть документ как HWPF (.doc): {Message}", e.Message);
                    // 'doc' останетс€ null, что приведет к попытке XWPF.
                }
                finally
                {
                    // Ётот блок гарантирует, что memoryStreamForHwpf будет утилизирован,
                    // *только если* его владение не было передано HWPFDocument.
                    // ≈сли memoryStreamForHwpf все еще содержит ссылку (т.е. HWPFDocument не был создан успешно),
                    // то мы должны его утилизировать.
                    memoryStreamForHwpf?.Dispose();
                    // nfs сам не IDisposable, поэтому не нужно вызывать Dispose().
                    // ≈сли nfs был создан, но не передан HWPFDocument, он будет собран сборщиком мусора,
                    // и его принадлежащий поток (memoryStreamForHwpf) будет утилизирован выше.
                }

                // --- ≈сли попытка HWPF не удалась, пробуем открыть как XWPF (.docx) ---
                if (doc == null)
                {
                    try
                    {
                        tmpStream.Position = 0; // —бросить оригинальный поток дл€ чтени€ XWPF
                        doc = new XWPFDocument(tmpStream); // XWPFDocument обычно Ќ≈ берет на себ€ владение tmpStream.
                                                           // ¬ызывающий код дл€ этого блока отвечает за утилизацию tmpStream.
                    }
                    catch (Exception e)
                    {
                        Logger?.Error("ќшибка при попытке открыть документ как XWPF (.docx): {Message}", e.Message);
                    }
                }

                // --- »тоговое присвоение ---
                if (doc != null)
                {
                    // ѕредполагаетс€, что 'Document' - это обертка, котора€ управл€ет жизненным циклом 'doc'
                    // (т.е. вызывает Dispose() на 'doc', если он IDisposable, как HWPFDocument).
                    Document = new(doc);
                    ExchangeValueFunc();
                }
            }
        }

        public virtual void InsertValue()
        {
            throw new NotImplementedException();
        }

        public virtual void ReadValue()
        {
            throw new NotImplementedException();
        }

        public virtual void UpdateValue()
        {
            throw new NotImplementedException();
        }
    }


    public class DataFrameView : WordExchange<DataFrame>
    {
        public DataFrameView(ExchangeOperation exchange, IProgress<int> progress = null)
            : base(exchange, progress) { }

        public override void ReadValue()
        {
            ExchangeValue = Document.DataFrames;
        }
    }


    public class TableView : WordExchange<TableValue>
    {
        public TableView(ExchangeOperation exchange, IProgress<int> progress = null)
            : base(exchange, progress) { }

        public override void ReadValue()
        {
            ExchangeValue = Document.Tables;
        }
    }

    public class ParagraphView : WordExchange<string>
    {
        public ParagraphView(ExchangeOperation exchange, IProgress<int> progress = null)
            : base(exchange, progress) { }

        public override void ReadValue()
        {
            ExchangeValue = Document.Paragraphs;
        }
    }
}
