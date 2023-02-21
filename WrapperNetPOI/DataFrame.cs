using Microsoft.Data.Analysis;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace WrapperNetPOI
{
    public class DataFrameView : ExchangeClass<DataFrame>
    {
        public int[] HeaderRows {set;get;}=new int[]{0};
        public string[] Header;
        
        public DataFrameView(ExchangeOperation exchangeType, string activeSheetName = "", DataFrame exchangeValue = null,
            IProgress<int> progress = null) : base(exchangeType, activeSheetName, progress) { }

        public override void ReadValue()
        {
            ReadValueHoleSheet();
        }

        protected internal void ReadHeader()
        {
            foreach (var head in HeaderRows)
            {
                for (int i = ActiveSheet.GetRow(head).FirstCellNum; i < ActiveSheet.GetRow(head).LastCellNum; i++)
                { 
                    if (Header == null)
                    {
                        Header = new string[ActiveSheet.GetRow(head).LastCellNum];
                        Header[i] = GetCellValue(ActiveSheet.GetRow(head).GetCell(i));
                    }
                    else
                    {
                        Header[i] =  $"{Header[i] ??""}{GetCellValue(ActiveSheet.GetRow(head).GetCell(i))}";
                    }
                }
            }

        }

        private void ReadValueHoleSheet() //Fast
        {
            ExchangeValue = new DataFrame();
            foreach (var column in Header)
            {
                DataFrameColumn dt = new StringDataFrameColumn(column);
                ExchangeValue.Columns.Add(dt);
            }


            List<string[]> tmpListString = new();

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
