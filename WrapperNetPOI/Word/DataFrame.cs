using Microsoft.Data.Analysis;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace WrapperNetPOI.Word
{
    public class DataFrameView : WordExchange<DataFrame>
    {
        public Header DataHeader { set; get; }

        public DataFrameView(ExchangeOperation exchangeType, string activeSheetName = "", DataFrame exchangeValue = null,
            Border border = null, IProgress<int> progress = null) : base(exchangeType, activeSheetName, border, progress) { }

        public override ISheet ActiveSheet
        {
            set
            {
                base.ActiveSheet = value;
                if (DataHeader == null)
                {
                    DataHeader = new Header
                    {
                        DFView = this,
                    };
                }
                else
                {
                    DataHeader.DFView = this;
                }
            }
            get
            {
                return base.ActiveSheet;
            }
        }

        public override void ReadValue()
        {
            ReadHeader();
            ReadValueHoleSheet();
        }

        public override void InsertValue()
        {
            base.InsertValue();
        }

        public void CreateHeader()
        {
            if (DataHeader != null)
            {
                if (WorkbookBorder != null)
                {
                    for (int i = WorkbookBorder.FirstColumn;
                        i < WorkbookBorder.LastColumn; i++)
                    {
                    }
                }
            }
        }

        protected void AppendOneRow(IRow row, DataFrame dataFrame)
        {
            ConvertType convert = new();
            List<KeyValuePair<string, object>> oneRow = new();
            foreach (var column in dataFrame.Columns)
            {
                ICell cell;
                var columnHeader = DataHeader.DataColumns.First(x => x.Name == column.Name);
                if (row != null)
                {
                    cell = row.GetCell(columnHeader.Number);
                }
                else
                {
                    cell = null;
                }
                var value = convert.GetValue(cell, column.DataType);
                oneRow.Add(new KeyValuePair<string, object>(columnHeader.Name, value));
            }
            dataFrame.Append(oneRow, true);
        }

        protected internal void ReadHeader()
        {
            DataHeader.GetHeaderRow();
            DataHeader.RenameDoubleHeaderColumn();
        }

        private void CreateColumns()
        {
            DataFrameColumn dt;
            foreach (var column in DataHeader.DataColumns)
            {
                switch (column.Type.Name)
                {
                    case "String":
                        dt = new StringDataFrameColumn(column.Name);
                        ExchangeValue.Columns.Add(dt);
                        break;

                    case "Double":
                        dt = new DoubleDataFrameColumn(column.Name);
                        ExchangeValue.Columns.Add(dt);
                        break;

                    case "DateTime":
                        dt = new DateTimeDataFrameColumn(column.Name);
                        ExchangeValue.Columns.Add(dt);
                        break;

                    default:
                        dt = new StringDataFrameColumn(column.Name);
                        ExchangeValue.Columns.Add(dt);
                        break;
                }
            }
        }

        private void ReadValueHoleSheet() //Fast
        {
            ExchangeValue = new DataFrame();
            CreateColumns();
            if (ActiveSheet != null)
            {
                int i = 0;
                foreach (IRow row in ActiveSheet)
                {
                    if (row.RowNum > i)
                    {
                        do
                        {
                            AppendOneRow(null, ExchangeValue);
                            i++;
                        }
                        while (row.RowNum != i);
                    }
                    if (!DataHeader.Rows.Contains(i))
                    {
                        if (WorkbookBorder == null)
                        {
                            AppendOneRow(row, ExchangeValue);
                        }
                        else if (WorkbookBorder != null)
                        {
                            if (i >= WorkbookBorder.FirstRow && i <= WorkbookBorder.LastRow)
                            {
                                AppendOneRow(row, ExchangeValue);
                            }
                            else if (i > WorkbookBorder.LastRow)
                            {
                                break;
                            }
                        }
                    }
                    i++;
                }
            }
        }
    }
}
