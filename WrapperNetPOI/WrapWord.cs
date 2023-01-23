/* ==================================================================
Copyright 2020-2022 sashamsuper

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
using NPOI.HWPF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.POIFS.Crypt;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
//using NPOI.SS.UserModel;
using Serilog;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace WrapperNetPOI
{

    public abstract class CellValue
    {
        string Text;
        ushort tableNumber;
        ushort rowNumber;
        ushort cellNumber;
        CellValue cellValue;
    }

    
    public abstract class DocumentWord
    {
        List<List<string[]>> Tables=new();
        private void XGetInformFromTable(IBody document)
        {
            List<string[]> rows = new();
            foreach (var table in document.Tables)
            {
                foreach (var row in table.Rows)
                {
                    string[] cells = default;
                    foreach (var cell in row.GetTableCells())
                    {
                        cells = row.GetTableCells().Select(x => x.GetText()).ToArray();
                    }
                    rows.Add(cells);
                }
                Tables.Add(rows);
            }
        }
        
        public virtual void GetTables()
        {
            return xDocument.Tables;
        }
        public virtual void GetParagraphs()
        {

        }
        private HWPFDocument hDocument;
        private XWPFDocument xDocument;
        public DocumentWord (object document)
        {
            if (document is HWPFDocument h)
            {
                hDocument=h;
            }
            else if (document is XWPFDocument x)
            {
                xDocument=x;
            }
        }


    }

    
    public class WrapperWord : Wrapper
    {

        public IDocument Document { set; get; }
        
        private XWPFDocument xDocument;
        private HWPFDocument hDocument;

        public WrapperWord(string pathToFile, IExchangeWord exchangeClass, ILogger logger = null) :
        base(pathToFile, exchangeClass, logger)
        { 

            //NPOI.HWPF.UserModel.Table
            Range range=hDocument.
            if  (range is NPOI.HWPF.UserModel.Table table)
            {

            } 
            
        }

        public void Exchange()
        {
            switch (exchangeClass.ExchangeOperationEnum)
            {
                case ExchangeOperation.Insert:
                    InsertValue();
                    break;
                case ExchangeOperation.Read:
                    ReadValue();
                    break;
                case ExchangeOperation.Update:
                    UpdateValue();
                    break;
                default:
                    Logger.Error("exchangeClass.ExchangeTypeEnum");
                    throw (new ArgumentOutOfRangeException("exchangeClass.ExchangeTypeEnum"));
            }
        }

        
        private void InsertValue()
        {
            if (File.Exists(PathToFile))
            {
                OnlyInsertValue();
            }
            else
            {
                CreateAndInsertValue();
            }
        }

        private void CreateAndInsertValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.InsertValue;
            ViewFile(FileMode.CreateNew, FileAccess.ReadWrite, true, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            //exchangeClass.Workbook.Write(fs, false);
            fs.Close();
        }

        private void ReadValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.ReadValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream, FileShare.Read);
        }

        private void UpdateValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.UpdateValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            //exchangeClass.Workbook.Write(fs, false);
            fs.Close();
        }

        private void OnlyInsertValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.InsertValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            //exchangeClass.Workbook.Write(fs, false);
            fs.Close();
        }


    }

    
}








