using NPOI.POIFS.Crypt;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.XWPF.UserModel;

namespace WrapperNetPOI
{
    public interface IExchangeWord:IExchange
    {
        
       
    }

    public class WordExchange:IExchangeWord
    {
        public WordExchange (ExchangeOperation exchange, string activeSheetName, IProgress<int> progress)
        {
            ExchangeOperationEnum = exchange;
            ProgressValue = progress;
        }

        public List<List<string[]>> Tables { set; get; } = new List<List<string[]>>();
        public IWorkbook Workbook { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public IProgress<int> ProgressValue { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public ILogger Logger { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public ExchangeOperation ExchangeOperationEnum { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public Action ExchangeValueFunc { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool CloseStream { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public void DeleteValue()
        {
            throw new NotImplementedException();
        }

        public void GetInternallyObject(Stream fs, bool addNew)
        {
            throw new NotImplementedException();
        }

        public void InsertValue()
        {
            throw new NotImplementedException();
        }

        public void ReadValue()
        {
            throw new NotImplementedException();
        }

        public void UpdateValue()
        {
            throw new NotImplementedException();
        }

        

    }
}