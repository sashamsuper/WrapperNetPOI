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
        public string[] Headers;
        
        public DataFrameView(ExchangeOperation exchangeType, string activeSheetName = "", DataFrame exchangeValue = null,
            IProgress<int> progress = null) : base(exchangeType, activeSheetName, progress) { }

        public override void ReadValue()
        {
            ReadValueHoleSheet();
        }

        private void ReadHeader()
        {
            foreach (var head in HeaderRows)
            {
                if (Headers==null)
                {
                    Headers=ActiveSheet.GetRow(head).Select(x => GetCellValue(x)).ToArray();
                }
                else
                {
                    var tmpHeaders=ActiveSheet.GetRow(head).Select(x => GetCellValue(x)).ToArray();
                    for (int i=0;i<Headers.Length;i++)
                    {
                        Headers[i]=Headers[i] +tmpHeaders.ElementAtOrDefault(i);
                    }
                }
            }

        }

        private void ReadValueHoleSheet() //Fast
        {

            DataFrameColumn[] columns = {
            //new StringDataFrameColumn("Name", names),
            //new PrimitiveDataFrameColumn<int>("Age", ages),
            //new PrimitiveDataFrameColumn<double>("Height", heights),
};

            ExchangeValue = new DataFrame();

            DataFrameColumn dt= new StringDataFrameColumn("Name");  
            //ExchangeValue.Columns.Add("sdd");




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
