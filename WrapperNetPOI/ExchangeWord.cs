using NPOI.SS.UserModel;
using NPOI.POIFS.Crypt;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using NPOI.XWPF.UserModel;

namespace WrapperNetPOI
{
    public interface IExchangeWord : IExchange
    {
    }

    public abstract class WordExchange : IExchangeWord
    {
        public WordExchange(ExchangeOperation exchange, IProgress<int> progress=null)
        {
            ExchangeOperationEnum = exchange;
            ProgressValue = progress;
        }

        public List<List<string[]>> Tables { set; get; } = new List<List<string[]>>();
        public IProgress<int> ProgressValue { get ; set ; }
        public ILogger Logger { get; set ; }
        public ExchangeOperation ExchangeOperationEnum { get; set; }
        public Action ExchangeValueFunc { get ; set ; }
        public List<TableValue> ExchangeValue { set; get; }
        public bool CloseStream { get ; set ; }

        public WordDoc Document {set;get;}

        public string Password {set;get;}

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
                 Document=new(doc);
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


    public class TableView : WordExchange
    {
        public TableView(ExchangeOperation exchange, IProgress<int> progress = null):
            base(exchange,progress)
        {}

        public override void ReadValue()
        {
            ExchangeValue = Document.GetTables();
        }

    }
}