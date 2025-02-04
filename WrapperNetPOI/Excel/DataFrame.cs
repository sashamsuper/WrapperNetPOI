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
using MathNet.Numerics.Optimization;
using Microsoft.Data.Analysis;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace System.Runtime.CompilerServices
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal class IsExternalInit { }
}


namespace WrapperNetPOI.Excel
{
    public static class Extensions
    {
        public static bool TryAddStandart<TKey, TValue>(
            this Dictionary<TKey, TValue> dictionary,
            TKey key,
            TValue value
        )
        {
            if (dictionary.ContainsKey(key))
            {
                return false;
            }
            dictionary.Add(key, value);
            return true;
        }

        public static bool TryAdd<TKey, TValue>(
            this Dictionary<TKey, TValue> dictionary,
            KeyValuePair<TKey, TValue> value
        )
        {
            return TryAddStandart(dictionary, value.Key, value.Value);
        }

         public static string ColumnNameFind(this DataFrame df,IEnumerable<string> findingColumnNames)
         {
            var findColumn = (
                from headerColumns in df.Columns
                join findingColums in findingColumnNames on headerColumns.Name equals findingColums
                select new { HeaderColumns = headerColumns, FindingColums = findingColums }
            )
                .FirstOrDefault()
                ?.FindingColums;
            return findColumn;
        }
        
    }

    public class Header
    {
        private int[] rows = { 0 };
        public int[] Rows
        {
            set { rows = value; }
            get
            {
                if (Border == null)
                {
                    return rows;
                }
                else
                {
                    List<int> headRows = new();
                    var tmpBorderList = Enumerable
                        .Range(Border.FirstRow, Border.LastRow + 1)
                        .ToList();
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
            private get { return dataFrameView; }
        }
        public Border Border { set; get; }

        public Header() { }

        public Header(int[] rows, Dictionary<int, Type> columns = null)
        {
            Rows = rows;
            if (columns != null)
            {
                CreateHeaderType(columns);
            }
        }

        public void CreateHeaderType(Dictionary<int, Type> columns)
        {
            List<DataColumn> tmp = new();
            foreach (var column in columns)
            {
                DataColumn columnHeader = new("", column.Key, column.Value);
                tmp.Add(columnHeader);
            }
            DataColumns = tmp.ToArray();
        }

        protected internal Type GetTypeOfCell(ISheet activeSheet, int columnNumber)
        {
            Dictionary<Type, int> conversionBall = new()
            {
                {typeof(String),0},
                {typeof(int),0},
                {typeof(Double),0},
                {typeof(DateTime),0}
            };
            //for (int i = 0; i < DataColumns.Length; i++)
            {
                //DataColumns[i] = new DataColumn("", i, typeof(String));
                for (int j = Border.FirstRow; j < Border.FirstRow + 10; j++)
                {
                    ICell cell = activeSheet.GetRow(j)?.GetCell(Border.FirstColumn + columnNumber);
                    WrapperCell wrapperCell = new(cell);
                    foreach (var x in conversionBall)
                    {
                        var value = wrapperCell.ToType(x.Key, wrapperCell.ThisCultureInfo);
                        if (wrapperCell.AutoType == x.Key)
                        {
                            conversionBall[x.Key]++;
                        }
                    }
                }
            }
            var valueType = conversionBall.OrderByDescending(x => x.Value).First().Key;
            return valueType;
        }

        protected internal virtual void GetNumberOfColumns(int rowsNumber)
        {
            {
                int countValue;
                /*if (Border.LastColumn != Border.FirstColumn)
                {
                    countValue = Border.LastColumn - Border.FirstColumn;
                }
                else
                */
                {
                    var lastColumn = 0;
                    if (Border.FirstRow != 0)
                    {
                        lastColumn = DFView.ActiveSheet.GetRow(Border.FirstRow).LastCellNum;
                    }
                    else
                    {
                        if (Rows.Length != 0)
                        {
                            lastColumn = DFView.ActiveSheet.GetRow(Rows[rowsNumber]).LastCellNum;
                        }
                        else
                        {
                            lastColumn = DFView.ActiveSheet.GetRow(Border.FirstRow).LastCellNum;
                        }
                    }
                    countValue = lastColumn - Border.FirstColumn;
                    DFView.WorkbookBorder.CorrectBorder(lastColumn: lastColumn);
                }
                if (DataColumns == null)
                {
                    DataColumns = new DataColumn[countValue];
                    for (int i = 0; i < DataColumns.Length; i++)
                    {
                        Type type = GetTypeOfCell(DFView.ActiveSheet, i);
                        //DataColumns[i] = new DataColumn("", i, typeof(String));
                        if (type != null)
                        {
                            DataColumns[i] = new DataColumn("", i, type);
                        }
                        else
                        {
                            DataColumns[i] = new DataColumn("", i, typeof(String));
                        }
                    }
                }
                for (int k = 0; k < DataColumns.Length; k++)
                {
                    DataColumns[k].Number = k + Border.FirstColumn;
                }
            }
        }

        protected internal virtual void GetColumnsName()
        {
            string[] tmpColName;
            tmpColName = new string[DataColumns.Length];
            foreach (var j in Rows)
            {
                for (int i = 0; i < DataColumns.Length; i++)
                {
                    ICell cell = DFView.ActiveSheet
                        .GetRow(j)
                        ?.GetCell(i + DFView.WorkbookBorder.FirstColumn);
                    string columnName;
                    if (cell?.IsMergedCell==true)
                    {
                        columnName = NewBaseType.GetFirstCellInMergedRegion(cell)?.ToString().Trim();
                    }
                    else
                    {
                        columnName = cell?.ToString().Trim();
                    }
                    //convertType.GetValue<string>(cell);
                    columnName ??= "";
                    if (tmpColName[i] != columnName)
                    {
                        tmpColName[i] = $"{tmpColName[i] ?? ""}{columnName}";
                    }
                }
            }
            for (int i = 0; i < DataColumns.Length; i++)
            {
                for (int j = 1; j < 15; j++)
                {
                    if (!DataColumns.Select(x => x.Name).Contains(tmpColName[i]))
                    {
                        DataColumns[i].Name = tmpColName[i].Trim();
                        break;
                    }
                    else
                    {
                        tmpColName[i] = tmpColName[i].Trim() + j;
                    }
                }
            }
        }

        protected internal virtual void GetHeaderRow()
        {
            if (Rows.Length == 0)
            {
                GetNumberOfColumns(0);
            }
            else
            {
                GetNumberOfColumns(Rows[0]);
            }
            GetColumnsName();
        }

        public void RenameDoubleHeaderColumn()
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

        public DataFrameView(
            ExchangeOperation exchangeType,
            string activeSheetName = "",
            DataFrame exchangeValue = null,
            Border border = null,
            Header header = null,
            IProgress<int> progress = null
        )
            : base(exchangeType, activeSheetName, border, progress)
        {
            ExchangeValue = exchangeValue;
            DataHeader = header;
        }

        public override ISheet ActiveSheet
        {
            set
            {
                base.ActiveSheet = value;
                if (DataHeader == null)
                {
                    DataHeader = new Header { DFView = this, };
                }
                else
                {
                    DataHeader.DFView = this;
                }
            }
            get { return base.ActiveSheet; }
        }

        public override void ReadValue()
        {
            ReadHeader();
            ReadValueHoleSheet();
        }

        public override void InsertValue()
        {
            if (DataHeader.Rows.Length != 0)
            {
                if (ExchangeValue != null)
                {
                    AddOneHeaderExcelRow(0);
                }
                WorkbookBorder.FirstRow = WorkbookBorder.FirstRow + 1;
            }
            for (int i =0; i < ExchangeValue.Rows.Count; i++)
            {
                AddOneExcelRow(i);
            }
        }
        private void AddOneHeaderExcelRow(int row)
        {
            int viewExcelRow = WorkbookBorder.Row(row);
            int columnsCount = ExchangeValue.Columns.Count;
            for (int j = 0; j < columnsCount; j++)
            {
                int viewExcelCol = WorkbookBorder.Column(j);
                Type dataType = ExchangeValue.Columns[j].DataType;
                IRow dataRow = ActiveSheet.GetRow(viewExcelRow) ?? ActiveSheet.CreateRow(viewExcelRow);
                CellType cellType = WrapperCell.ReturnCellType(dataType);
                ICell cell = dataRow.GetCell(viewExcelCol) ?? dataRow.CreateCell(viewExcelCol, cellType);
                var value = ExchangeValue.Columns[j].Name;
                WrapperCell wrapperCell = new(cell);
                wrapperCell.SetValue(value);
            }
        }

    

    private void AddOneExcelRow(int row)
        {
            int viewExcelRow = WorkbookBorder.Row(row);
            for (int j = 0; j < ExchangeValue.Columns.Count; j++)
            {
                int viewExcelCol = WorkbookBorder.Column(j);
                Type dataType = ExchangeValue.Columns[j].DataType;
                IRow dataRow =
                    ActiveSheet.GetRow(viewExcelRow) ?? ActiveSheet.CreateRow(viewExcelRow);
                CellType cellType = WrapperCell.ReturnCellType(dataType);
                ICell cell =
                    dataRow.GetCell(viewExcelCol) ?? dataRow.CreateCell(viewExcelCol, cellType);
                object value;
                if (ExchangeValue.Rows[row][j]==null)
                {
                    if (dataType==typeof(String))
                    {
                        value="";
                    }
                    else
                    {
                        value=Activator.CreateInstance(dataType);
                    }
                }
                else
                {
                    value = Convert.ChangeType(ExchangeValue.Rows[row][j], dataType);
                }
                WrapperCell wrapperCell = new(cell);
                wrapperCell.SetValue(value);
            }
        }

        protected void AppendOneRow(IRow row, DataFrame dataFrame)
        {
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
                var value = new WrapperCell(cell).GetValue(column.DataType);
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
                    case "Int32":
                        dt = new Int32DataFrameColumn(column.Name);
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
                    case "Boolean":
                        dt = new BooleanDataFrameColumn(column.Name);
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
                        } while (row.RowNum != i);
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
