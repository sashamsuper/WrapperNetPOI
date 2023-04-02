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

using Microsoft.Data.Analysis;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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

    public class Header
    {
        private int[] rows = new int[] { 0 };

        public int[] Rows
        {
            init
            {
                rows = value;
            }
            get
            {
                if (Border == null)
                {
                    return rows;
                }
                else
                {
                    List<int> headRows = new();
                    var tmpBorderList = Enumerable.Range(Border.FirstRow, Border.LastRow + 1).ToList();
                    foreach (var x in rows)
                    {
                        headRows.Add(tmpBorderList[x]);
                    }
                    return headRows.ToArray();
                }
            }
        }

        public DataColumn[] DataColumns { set; get; }
        private DataFrameView dataFrameView;

        public DataFrameView DFView
        {
            set
            {
                dataFrameView = value;
                Border = dataFrameView.WorkbookBorder;
            }
            private get
            {
                return dataFrameView;
            }
        }

        public Border Border { set; get; }

        public void CreateHeaderType(Dictionary<int, Type> columns)
        {
            List<DataColumn> tmp = new();

            foreach (var column in columns)
            {
                DataColumn columnHeader =
                new("", column.Key, column.Value);
                tmp.Add(columnHeader);
            }
            DataColumns = tmp.ToArray();
        }

        protected internal virtual void GetHeaderRow()
        {
            foreach (var j in Rows)
            {
                ConvertType convertType = new();
                if (j == Rows[0])
                {
                    int countValue;
                    if (Border.LastColumn != Border.FirstColumn)
                    {
                        countValue = Border.LastColumn - Border.FirstColumn;
                    }
                    else
                    {
                        var lastColumn = DFView.ActiveSheet.GetRow(Rows[j]).LastCellNum;
                        countValue = lastColumn - Border.FirstColumn;
                        DFView.WorkbookBorder.
                               CorrectBorder(lastColumn: lastColumn);
                    }
                    if (DataColumns == null)
                    {
                        DataColumns = new DataColumn[countValue];
                        for (int i = 0; i < DataColumns.Length; i++)
                        {
                            DataColumns[i] = new DataColumn("", i, typeof(String));
                        }
                    }
                    for (int k = 0; k < DataColumns.Length; k++)
                    {
                        DataColumns[k].Number = k + Border.FirstColumn;
                    }
                }
                for (int i = 0; i < DataColumns.Length; i++)
                {
                    ICell cell = DFView.ActiveSheet.GetRow(j)?.GetCell(i + DFView.WorkbookBorder.FirstColumn);
                    string columnName = convertType.GetValue<string>(cell);
                    columnName ??= "";
                    DataColumns[i].Name = $"{DataColumns[i].Name ?? ""}{columnName}";
                }
            }
        }

        public void RenameDobleHeaderColumn()
        {
            for (int i = DataColumns.Length - 1; i >= 0; i--)
            {
                int j = 0;
                string tmpHeader = DataColumns[i].Name;
                while (DataColumns.Count(x => x.Name == DataColumns[i].Name) > 1)
                {
                    j++;
                    DataColumns[i].Name = $"{tmpHeader}{j}";
                }
            }
        }
    }

    public class DataColumn
    {
        public string Name { set; get; }
        public int Number { set; get; }
        public Type Type { set; get; }

        public override string ToString()
        {
            return Name;
        }

        public DataColumn(string name, int columnNumber, Type columnType)
        {
            Name = name;
            Number = columnNumber;
            Type = columnType;
        }
    }

    public class DataFrameView : ExchangeClass<DataFrame>
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
            DataHeader.RenameDobleHeaderColumn();
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
#if DEBUG
                    if (i % 1000 == 0)
                    {
                        Debug.WriteLine(i);
                    }
#endif
                }
            }
        }
    }
}