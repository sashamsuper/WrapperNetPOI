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
using NPOI.POIFS.Crypt;
using NPOI.XWPF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.HSSF.UserModel;


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
                object doc=null;
                POIFSFileSystem nfs=default;
                try
                {
                    MemoryStream memoryStream = new MemoryStream();
                    tmpStream.CopyTo(memoryStream); // Ensure the stream is reset to the beginning
                    memoryStream.Position = 0;
                    nfs = new(memoryStream);
                }
                catch (Exception e)
                {
                    // If the file is not a POIFSFileSystem, it might be an XWPF document
                    Logger?.Error("Error reading Word document: {Message}", e.Message);
                }
                if (nfs!=default && nfs.Root.HasEntry("WordDocument"))
                {
                    doc = new HWPFDocument(nfs);
                }
                else
                {
                    try
                    {
                        tmpStream.Position = 0; // Reset the original stream position
                        doc = new XWPFDocument(tmpStream);
                    }
                    catch (Exception e)
                    {
                        Logger?.Error("Error reading Word XWPF document: {Message}", e.Message);
                    }
                }

                if (doc != null)
                {
                    Document = new(doc);
                }
            }
            //exchangeClass.ActiveSheet = ActiveSheet;
            if (Document != null)
            {
                ExchangeValueFunc();
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
