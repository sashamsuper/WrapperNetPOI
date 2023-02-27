using Microsoft.Data.Analysis;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace WrapperNetPOI
{

    public static class Extensions
    {
        public static bool TryAddStandart<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue value)
        {
            if (dictionary.ContainsKey(key))
            {
                return false;
            }
            dictionary.Add(key, value);
            return true;
        }
        public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, KeyValuePair<TKey, TValue> value)
        {
            return TryAddStandart(dictionary, value.Key, value.Value);
        }
    }


    public class ColumnHeader
    {
        public string Name {set;get;}
        public int ColumnNumber { set; get; }
        public Type ColumnType { set; get; }

        public override string ToString()
        {
            return Name;
        }

    }

    
    
    public class DataFrameView : ExchangeClass<DataFrame>
    {
        
        
        public int[] HeaderRows {set;get;}=new int[]{0};
        public ColumnHeader[] Header { set; get; } 
        
        public DataFrameView(ExchangeOperation exchangeType, string activeSheetName = "", DataFrame exchangeValue = null,
            IProgress<int> progress = null) : base(exchangeType, activeSheetName, progress) 
            {}

        public override void ReadValue()
        {
            ReadHeader();
            ReadValueHoleSheet();
        }





        protected void GetOneRow(IRow row, DataFrame dataFrame, DataFrameRow frameRow)
        {
            foreach (var column in dataFrame.Columns)
            {
                ConvertType convert = new ();
                var columnHeader=Header.Where(x => x.Name == column.Name).First();
                ICell cell = row.GetCell(columnHeader.ColumnNumber);
                var value=convert.GetValue(cell, column.DataType);
                frameRow[columnHeader.ColumnNumber] = value;
            }

        }

        protected internal void ReadHeader()
        {
            Header = default;
            foreach (var head in HeaderRows)
            {
                var counValue = ActiveSheet.GetRow(head).LastCellNum -
                                ActiveSheet.GetRow(head).FirstCellNum; // +1 ruled out, NPOI feature


                if (Header == null)
                {
                    Header = new ColumnHeader[counValue];
                    for (int j = 0; j < Header.Length; j++)
                    {
                        Header[j] = new ColumnHeader 
                        { 
                            ColumnNumber= j
                        };
                    }
                }
                
                for (int i = 0; i < Header.Length; i++)
                {
                    ConvertType convertType = new();
                    string columnName;
                    columnName = convertType.GetValue<string>(ActiveSheet.GetRow(head).GetCell(i));
                    columnName ??= "";
                    Header[i].Name= $"{Header[i].Name ?? ""}{columnName}";
                }
            }
            for (int i= Header.Length-1; i>=0;i--)
            {
                int j = 0;
                string tmpHeader = Header[i].Name;
                while (Header.Count(x => x == Header[i]) > 1)
                {
                    j++;
                    Header[i].Name = $"{tmpHeader}{j}";
                }
            }
        }

        private void ReadValueHoleSheet() //Fast
        {
            ExchangeValue = new DataFrame();
            foreach (var column in Header)
            {
                DataFrameColumn dt = new StringDataFrameColumn(column.Name);
                ExchangeValue.Columns.Add(dt);
            }

            if (ActiveSheet != null)
            {
                int i = 0;
                foreach (IRow value in ActiveSheet)
                {
                    IRow row = value;
                    if (row.RowNum > i)
                    {
                        while (true)
                        {
                            tmpListString.Add(Array.Empty<string>());
                            i++;
                            if (row.RowNum == i)
                            {
                                break;
                            }
                        }
                    }
                    tmpListString.Add(row.Select(x => GetCellValue(x)).ToArray());
                    i++;
#if DEBUG
                    if (i % 1000 == 0)
                    {

                        Debug.WriteLine(i);

                    }
#endif
                }
                //ExchangeValue = tmpListString;
            }
        }


    }
}
