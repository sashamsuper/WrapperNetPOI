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
using System.ComponentModel;
using System.Linq;

namespace System.Runtime.CompilerServices
{
    /// <summary>
    /// The checks if is external init.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal class IsExternalInit { }
}

namespace WrapperNetPOI.Excel
{
    /// <summary>
    /// The extensions.
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Try add standart.
        /// </summary>
        /// <typeparam name="TKey"/>
        /// <typeparam name="TValue"/>
        /// <param name="dictionary">The dictionary.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <returns>A bool</returns>
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

        /// <summary>
        /// Try add.
        /// </summary>
        /// <typeparam name="TKey"/>
        /// <typeparam name="TValue"/>
        /// <param name="dictionary">The dictionary.</param>
        /// <param name="value">The value.</param>
        /// <returns>A bool</returns>
        public static bool TryAdd<TKey, TValue>(
            this Dictionary<TKey, TValue> dictionary,
            KeyValuePair<TKey, TValue> value
        )
        {
            return TryAddStandart(dictionary, value.Key, value.Value);
        }

        /// <summary>
        /// Column name find.
        /// </summary>
        /// <param name="df">The df.</param>
        /// <param name="findingColumnNames">The finding column names.</param>
        /// <returns>A string</returns>
        public static string ColumnNameFind(
            this DataFrame df,
            IEnumerable<string> findingColumnNames
        )
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

        /// <summary>
        /// Add the column.
        /// </summary>
        /// <typeparam name="T"/>
        /// <param name="df">The df.</param>
        /// <param name="name">The name.</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NotSupportedException"></exception>
        public static void AddColumn<T>(this DataFrame df, string name)
        {
            if (df == null) throw new ArgumentNullException(nameof(df));
            DataFrameColumn column;
            long count = df.Rows.Count;

            var t = typeof(T);

            if (t == typeof(string))
            {
                column = new StringDataFrameColumn(name, count);
            }
            else if (t == typeof(int))
            {
                column = new Int32DataFrameColumn(name, count);
            }
            else if (t == typeof(long))
            {
                column = new Int64DataFrameColumn(name, count);
            }
            else if (t == typeof(short))
            {
                column = new Int16DataFrameColumn(name, count);
            }
            else if (t == typeof(byte))
            {
                column = new ByteDataFrameColumn(name, count);
            }
            else if (t == typeof(sbyte))
            {
                column = new SByteDataFrameColumn(name, count);
            }
            else if (t == typeof(float))
            {
                column = new SingleDataFrameColumn(name, count);
            }
            else if (t == typeof(double))
            {
                column = new DoubleDataFrameColumn(name, count);
            }
            else if (t == typeof(decimal))
            {
                column = new DecimalDataFrameColumn(name, count);
            }
            else if (t == typeof(bool))
            {
                column = new BooleanDataFrameColumn(name, count);
            }
            else if (t == typeof(DateTime))
            {
                column = new DateTimeDataFrameColumn(name, count);
            }
            else
            {
                throw new NotSupportedException($"Type {t} is not supported.");
                // Для прочих типов используем универсальный примитивный столбец
                //column = new PrimitiveDataFrameColumn<T>(name, count);
            }
            df.Columns.Add(column);
        }

        /// <summary>
        /// Add V buffer column.
        /// </summary>
        /// <typeparam name="T"/>
        /// <param name="df">The df.</param>
        /// <param name="name">The name.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void AddVBufferColumn<T>(this DataFrame df, string name)
        {
            if (df == null) throw new ArgumentNullException(nameof(df));
            DataFrameColumn column;
            long count = df.Rows.Count;
            column = new VBufferDataFrameColumn<T>(name, count);
            df.Columns.Add(column);
        }
    }







    /// <summary>
    /// The header.
    /// </summary>
    public class Header : IHeader
    {
        /// <summary>
        /// The rows.
        /// </summary>
        private int[] rows = { 0 };
        /// <summary>
        /// Gets or sets the rows.
        /// </summary>
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
        /// <summary>
        /// Gets or sets the data columns.
        /// </summary>
        public DataColumn[] DataColumns { set; get; }
        /// <summary>
        /// The data frame view.
        /// </summary>
        private DataFrameView dataFrameView;
        /// <summary>
        /// Gets or sets the DF view.
        /// </summary>
        internal DataFrameView DFView
        {
            set
            {
                dataFrameView = value;
                Border = dataFrameView.WorkbookBorder;
            }
            private get { return dataFrameView; }
        }
        /// <summary>
        /// Gets or sets the border.
        /// </summary>
        internal Border Border { set; get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Header"/> class.
        /// </summary>
        public Header() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Header"/> class.
        /// </summary>
        /// <param name="rows">The rows.</param>
        /// <param name="columns">The columns.</param>
        public Header(int[] rows, Dictionary<int, Type> columns = null)
        {
            Rows = rows;
            if (columns != null)
            {
                CreateHeaderType(columns);
            }
        }

        /// <summary>
        /// Creates header type.
        /// </summary>
        /// <param name="columns">The columns.</param>
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

        /// <summary>
        /// Get type of cell.
        /// </summary>
        /// <param name="activeSheet">The active sheet.</param>
        /// <param name="columnNumber">The column number.</param>
        /// <returns>A Type</returns>
        protected internal Type GetTypeOfCell(ISheet activeSheet, int columnNumber)
        {
            Dictionary<Type, int> conversionBall =
                new()
                {
                    { typeof(String), 0 },
                    { typeof(int), 0 },
                    { typeof(Double), 0 },
                    { typeof(DateTime), 0 }
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

        /// <summary>
        /// Get number of columns.
        /// </summary>
        /// <param name="rowsNumber">The rows number.</param>
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
                            var row = DFView.ActiveSheet.GetRow(Rows[rowsNumber]);
                            if (row != null)
                                lastColumn = DFView.ActiveSheet.GetRow(Rows[rowsNumber]).LastCellNum;
                            else
                                lastColumn = 0;
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

        /// <summary>
        /// Get columns name.
        /// </summary>
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
                    if (cell?.IsMergedCell == true)
                    {
                        columnName = NewBaseType
                            .GetFirstCellInMergedRegion(cell)
                            ?.ToString()
                            .Trim();
                    }
                    else
                    {
                        columnName = cell?.ToString().Trim();
                    }
                    //convertType.GetValue<string>(cell);
                    columnName ??= "";
                    if (tmpColName[i] != columnName)
                    {
                        tmpColName[i] = $"{tmpColName[i] ?? ""}{columnName}".Trim();
                    }
                }
            }
            for (int i = 0; i < DataColumns.Length; i++)
            {
                if (String.IsNullOrWhiteSpace(tmpColName[i]))
                {
                    tmpColName[i] = "_";
                }
                var constNameValue = tmpColName[i];
                for (int j = 1; j < 15; j++)
                {
                    if (!DataColumns.Select(x => x.Name).Contains(tmpColName[i]))
                    {
                        DataColumns[i].Name = tmpColName[i];
                        break;
                    }
                    else
                    {
                        tmpColName[i] = $"constNameValue{j}";
                    }
                }
            }
        }

        /// <summary>
        /// Get header row.
        /// </summary>
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

        /// <summary>
        /// Rename double header column.
        /// </summary>
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



    /// <summary>
    /// The data frame view.
    /// </summary>
    public class DataFrameView : ExchangeClass<DataFrame>, IDataFrameView
    {
        /// <summary>
        /// Gets or sets the data header.
        /// </summary>
        public Header DataHeader { set; get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataFrameView"/> class.
        /// </summary>
        /// <param name="exchangeType">The exchange type.</param>
        /// <param name="activeSheetName">The active sheet name.</param>
        /// <param name="exchangeValue">The exchange value.</param>
        /// <param name="border">The border.</param>
        /// <param name="header">The header.</param>
        /// <param name="progress">The progress.</param>
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

        /// <summary>
        /// Gets or sets the active sheet.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the data header.
        /// </summary>
        IHeader IDataFrameView.DataHeader { get; set; }

        /// <summary>
        /// Reads the value.
        /// </summary>
        public override void ReadValue()
        {
            ReadHeader();
            ReadValueHoleSheet();
        }

        /// <summary>
        /// Inserts the value.
        /// </summary>
        public override void UpdateValue()
        {
            _UpdateValue();
        }

        private void _UpdateValue(bool addHeader = true)
        {
            if (DataHeader.Rows.Length != 0)
            {
                if (addHeader)
                {
                    AddOneHeaderExcelRow(0);
                    WorkbookBorder.FirstRow = WorkbookBorder.FirstRow + 1;
                }
            }
            for (int i = 0; i < ExchangeValue.Rows.Count; i++)
            {
                AddOneExcelRow(i);
            }
        }

        public override void InsertValue()
        {
            int rowsCount = ActiveSheet.RowsCount();
            WorkbookBorder.FirstRow = rowsCount;
            if (rowsCount > 0)
            {
                _UpdateValue(false);
            }
            else
            {
                _UpdateValue();
            }
        }



        /// <summary>
        /// Add one header excel row.
        /// </summary>
        /// <param name="row">The row.</param>
        private void AddOneHeaderExcelRow(int row)
        {
            int viewExcelRow = WorkbookBorder.Row(row);
            int columnsCount = ExchangeValue.Columns.Count;
            for (int j = 0; j < columnsCount; j++)
            {
                int viewExcelCol = WorkbookBorder.Column(j);
                Type dataType = ExchangeValue.Columns[j].DataType;
                IRow dataRow =
                    ActiveSheet.GetRow(viewExcelRow) ?? ActiveSheet.CreateRow(viewExcelRow);
                CellType cellType = WrapperCell.ReturnCellType(dataType);
                ICell cell =
                    dataRow.GetCell(viewExcelCol) ?? dataRow.CreateCell(viewExcelCol, cellType);
                var value = ExchangeValue.Columns[j].Name;
                WrapperCell wrapperCell = new(cell);
                wrapperCell.SetValue(value);
            }
        }

        /// <summary>
        /// Add one excel row.
        /// </summary>
        /// <param name="row">The row.</param>
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
                if (ExchangeValue.Rows[row][j] == null)
                {
                    if (dataType == typeof(String))
                    {
                        value = "";
                    }
                    else
                    {
                        value = Activator.CreateInstance(dataType);
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

        /// <summary>
        /// Appends one row.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="dataFrame">The data frame.</param>
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

        /// <summary>
        /// Reads the header.
        /// </summary>
        public void ReadHeader()
        {
            DataHeader.GetHeaderRow();
            DataHeader.RenameDoubleHeaderColumn();
        }

        /// <summary>
        /// Creates the columns.
        /// </summary>
        public void CreateColumns()
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

        /// <summary>
        /// Reads value hole sheet.
        /// </summary>
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
