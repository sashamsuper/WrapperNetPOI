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
using System.Diagnostics;
using SixLabors.ImageSharp.ColorSpaces;
using Serilog;

/* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
До:
using NPOI.POIFS.Crypt;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
//using NPOI.SS.UserModel;
using Serilog;
using System;
using System.Diagnostics;
После:
using NPOI.XWPF.Crypt;
using NPOI.SS.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
//using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
*/
//using NPOI.SS.UserModel;
//using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace WrapperNetPOI
{

    public class CellValue
    {
        public string text;
        public int tableNumber;
        public int rowNumber;
        public int cellNumber;
        public int level;
        
        public CellValue(string text, int tableNumber, int rowNumber, int cellNumber,int level)
        {
            this.text = text;
            this.tableNumber= tableNumber;
            this.rowNumber= rowNumber;
            this.cellNumber= cellNumber;
            this.level= level;
    }
    }


    public class WordDoc
    {
        public List<CellValue> cells;

        private List<CellValue> XGetTables(IBody body, int level=0)
        {
            List<CellValue> cells=new();
            int i = 0; int j = 0; int k = 0;
            foreach (XWPFTable table in body.Tables)
            {
                i++;
                foreach (XWPFTableRow row in table.Rows)
                {
                    j++;
                    foreach (XWPFTableCell cell in row.GetTableICells())
                    {
                        k++;
                        if (cell?.BodyElements.Count > 0)
                        {
                            CellValue cellValue = new(cell.GetText(), i, j, k, level+1);
                            XGetTables(cell);
                            cells.Add(cellValue);
                        }
                        else
                        {
                            CellValue cellValue = new(cell.GetText(), i, j, k, level);
                            cells.Add(cellValue);
                        }
                    }
                }
            }
            return cells;
        }


        public virtual void GetTables()
        {
            if (Document is XWPFDocument x)
            {
               
                //foreach (var y in x.BodyElements)
                {
                    cells = XGetTables(x);
                }
            }
        }
        public virtual void GetParagraphs()
        {

        }
        private HWPFDocument hDocument;
        private XWPFDocument xDocument;
        public object Document
        {
            set
            { if (value is HWPFDocument h)
                {
                    hDocument = h;
                }
                else if (value is XWPFDocument x)
                {
                    xDocument = x;
                }
            }
            get 
            {
                if (xDocument is null)
                {
                    return hDocument;
                }
                else
                { 
                    return xDocument;
                }
            }
        }


    }


    public class WrapperWord : Wrapper
    {

        public WordDoc Document { set; get; }

        public WrapperWord(string pathToFile, IExchangeWord exchangeClass, ILogger logger = null) :
        base(pathToFile, exchangeClass, logger)
        {

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








