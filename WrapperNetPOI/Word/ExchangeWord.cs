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

using NPOI.POIFS.Crypt;
using NPOI.XWPF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;

namespace WrapperNetPOI.Word
{
    public interface IExchangeWord : IExchange
    {
    }

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
        public bool CloseStream { get; set; }

        public WordDoc Document { set; get; }

        public string Password { set; get; }

        public void DeleteValue()
        {
            throw new NotImplementedException();
        }

        public void GetInternallyObject(Stream tmpStream, bool addNew)
        {
            FileStream fs = default;
            if (Password != null)
            {
                NPOI.POIFS.FileSystem.POIFSFileSystem nfs =
                new(fs);
                EncryptionInfo info = new(nfs);
                Decryptor dc = Decryptor.GetInstance(info);
                //bool b = dc.VerifyPassword(Password);
                dc.VerifyPassword(Password);
                tmpStream = dc.GetDataStream(nfs);
            }
            if (addNew)
            {
                /*
                Workbook = new XSSFWorkbook();
                Workbook.CreateSheet(ActiveSheetName);
                Ac
                tiveSheet = Workbook.GetSheet(ActiveSheetName);
                */
            }
            else
            {
                XWPFDocument doc = new(tmpStream);
                Document = new(doc);
            }
            //exchangeClass.ActiveSheet = ActiveSheet;
            ExchangeValueFunc();
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
        public TableView(ExchangeOperation exchange, IProgress<int> progress = null) :
            base(exchange, progress)
        { }

        public override void ReadValue()
        {
            ExchangeValue = Document.Tables;
        }
    }

    public class ParagraphView : WordExchange<string>
    {
        public ParagraphView(ExchangeOperation exchange, IProgress<int> progress = null) :
            base(exchange, progress)
        { }

        public override void ReadValue()
        {
            ExchangeValue = Document.Paragraphs;
        }
    }
}