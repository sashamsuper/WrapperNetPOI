﻿/* ==================================================================
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

        public CellValue(string text, int tableNumber, int rowNumber, int cellNumber, int level)
        {
            this.text = text;
            this.tableNumber = tableNumber;
            this.rowNumber = rowNumber;
            this.cellNumber = cellNumber;
            this.level = level;
        }
    }


    public class WordDoc
    {
        private List<CellValue> cells;

        private List<CellValue> XGetTables(IBody body, ref int tableN,int level = 0)
        {
            List<CellValue> cells = new();
            int i = tableN; int j = 0; int k = 0;
            foreach (XWPFTable table in body.Tables)
            {
                foreach (XWPFTableRow row in table.Rows)
                {
                    foreach (XWPFTableCell cell in row.GetTableICells())
                    {
                        if (cell?.BodyElements.Count > 0)
                        {
                            CellValue cellValue = new(cell.GetText(), i, j, k, level);
                            cells.Add(cellValue);
                            cells.AddRange(XGetTables(cell,ref i,level+1));
                        }
                        else
                        {
                            CellValue cellValue = new(cell.GetText(), i, j, k, level);
                            cells.Add(cellValue);
                        }
                        k++;
                    }
                    j++;
                }
                i++;
            }
            return cells;
        }


        public virtual List<CellValue> GetCells()
        {
            if (Document1 is XWPFDocument x)
            {
                int i=0;
                cells = XGetTables(x,ref i);
                return cells;
            }
            return default;
        }

        public class TableValue
        {
            public int tableNumber;
            public int level;

            public List<string[]> Value;

            public TableValue(int tableNumber,int level)
            {
                this.tableNumber = tableNumber;
                this.level = level;
            }
        }


        public virtual List<CellValue> GetTables()
        {
            var tables=cells.GroupBy(t=>t.tableNumber).
            Select(table=>table.GroupBy(r=>r.rowNumber).OrderBy(rowN=>rowN.Key).Select(row=>row.OrderBy(cell=>cell.cellNumber).ToArray()));

            var tables2=cells.GroupBy(t=>t.tableNumber).
            Select(table=>table.GroupBy(r=>r.rowNumber).OrderBy(rowN=>rowN.Key).Select(row=>
            new {tableNumber=table.Key,rowNumber=row.Key, value=row.OrderBy(cell=>cell.cellNumber).
            Select(str=>str.text).ToArray()}));
            
            List<TableValue> tableList=new();
            foreach (var table in tables2)
            {
                TableValue tableV = new(table.First().tableNumber,table.First().tableNumber);
                List<string[]> rows=new();
                foreach (var row in table)
                {
                     rows.Add(row.value);
                }
                tableV.Value=rows;
                tableList.Add(tableV);
            }


            if (Document1 is XWPFDocument x)
            {
                int i=0;
                cells = XGetTables(x,ref i);
                return cells;
            }
            return default;
        }

        public virtual void GetParagraphs()
        {

        }

        public WordDoc(object doc)
        {
            this.Document1 = doc;
        }

        private HWPFDocument hDocument;
        private XWPFDocument xDocument;
        public object Document1
        {
            set
            {
                if (value is HWPFDocument)
                {
                    HWPFDocument hDocument = new();
                }
                else if (value is XWPFDocument x)
                {
                    xDocument = x;
                }
            }
            get
            {
                if (hDocument is null)
                {
                    return xDocument;
                }
                else
                {
                    return hDocument;
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

        protected override void InsertValue()
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

        protected override void ReadValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.ReadValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream, FileShare.Read);
        }

        protected override void UpdateValue()
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







