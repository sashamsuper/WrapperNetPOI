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

using NPOI.POIFS.Crypt;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WrapperNetPOI
{

    public enum ExchangeOperation
    {
        Insert,
        Read,
        Update,
        Delete
    }

    class Border // in developing
    {
        
        int FirstViewedRow { set; get; }
        int LastViewedRow { set; get; }
        int FirstViewedColumn { set; get; }
        int LastViewedColumn { set; get; }
    }

    public static class Extension
    {
        public static int RowsCount(this ISheet sheet)
        {
            if (sheet != null)
            {
                if (sheet.LastRowNum == sheet.FirstRowNum)
                {
                    return 0;
                }
                else
                {
                    return sheet.LastRowNum - sheet.FirstRowNum + 1;
                }
            }
            else
            {
                return 0;
            }
        }

        public static IList<string[]> ConvnertToMatrix(IList<string> list)
        {
            return list?.Select(x => new string[] { x }).ToList();
        }

        public static IList<string[]> ConvnertToMatrix(IDictionary<string, string[]> dict)
        {
            return dict?.SelectMany(x => x.Value.ToList(), (key, value) => new string[] { key.Key, value }).ToList();
        }

        public static IList<string> ConvertToList(IList<string[]> list)
        {
            return list.Select(x => String.Join("", x)).ToList();
        }

        public static IDictionary<string, string[]> ConvertToDictionary(IList<string[]> list)
        {
            Dictionary<string, string[]> dict = new();
            var groupValue = list.Where(x => x[0] != null).GroupBy(y => y[0], (value) => value[1]);
            foreach (var group in groupValue)
            {
                dict[group.Key] = group.Select(x => x).ToArray();
            }
            return dict;
        }
    }

    public interface IExchange
    {
        IProgress<int> ProgressValue { set; get; }
        ILogger Logger { set; get; }
        string ActiveSheetName { set; get; }
        ExchangeOperation ExchangeOperationEnum { set; get; }
        int FirstViewedRow { set; get; }
        int LastViewedRow { set; get; }
        int FirstViewedColumn { set; get; }
        int LastViewedColumn { set; get; }
        ISheet ActiveSheet { set; get; }
        Action ExchangeValueFunc { set; get; }
        bool CloseStream { get; set; }
        void ReadValue();
        void InsertValue();
        void UpdateValue();
        void DeleteValue();
    }

    public abstract class ExchangeClass<Tout> : IExchange
    {
        public virtual bool CloseStream { set; get; } = true;

        public IProgress<int> ProgressValue { set; get; }

        public static int ReturnProgress(int number, int total)
        {
            if (total != 0)
            {
                return ((number * 100) / (total));
            }
            else
            {
                return 0;
            }
        }

        public ILogger Logger { set; get; }
        public ExchangeClass(ExchangeOperation exchange, string activeSheetName, IProgress<int> progress)
        {
            ExchangeOperationEnum = exchange;
            ActiveSheetName = activeSheetName;
            ProgressValue = progress;
        }

        public string ActiveSheetName { set; get; }

        public ExchangeOperation ExchangeOperationEnum { set; get; }
        public ISheet ActiveSheet { set; get; }
        public int FirstViewedRow
        {
            set
            {
                firstViewedRow = value;
            }
            get
            {
                if (firstViewedRow == null)
                {

                    return (ActiveSheet != null) ? ActiveSheet.FirstRowNum : 0;
                }
                else
                {
                    return firstViewedRow ?? 0;
                }
            }
        }
        private int? firstViewedRow;
        public int FirstViewedColumn
        {
            set
            {
                firstViewedColumn = value;
            }
            get
            {
                return firstViewedColumn ?? 0;
            }
        }
        private int? firstViewedColumn;
        public int LastViewedRow
        {
            set
            {
                lastViewedRow = value;
            }
            get
            {
                if (lastViewedRow == null || lastViewedRow == 0)
                {
                    return (ActiveSheet != null) ? ActiveSheet.LastRowNum : 0;
                }
                else
                {
                    return lastViewedRow ?? 0;
                }
            }
        }
        private int? lastViewedRow;
        public int LastViewedColumn
        {
            set
            {
                lastViewedColumn = value;
            }
            get
            {
                return lastViewedColumn ?? 0;
            }
        }
        private int? lastViewedColumn;
        public Tout ExchangeValue { set; get; }
        public Action ExchangeValueFunc { set; get; }

        /// <summary>
        /// $Return Date by dd.mm.yyyy$
        /// </summary>
        /// <param name="date">The date<see cref="DateTime"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string ReturnStringDate(DateTime date)
        {
            var day = String.Format("{0:D2}", date.Day);
            var mounth = String.Format("{0:D2}", date.Month);
            var year = String.Format("{0:D4}", date.Year);
            return $"{day}.{mounth}.{year}";
        }

        public string GetCellValue(ICell cell)
        {
            try
            {
                string returnValue = default;
                if (cell != null)
                {
                    if (cell.IsMergedCell)
                    {
                        cell = GetFirstCellInMergedRegion(cell);
                    }
                    if (cell?.CellType == CellType.Numeric
                      && cell.NumericCellValue > 36526 &&
                      cell.NumericCellValue < 47484)
                    {
                        returnValue = ReturnStringDate(cell.DateCellValue);
                    }
                    else if (cell?.CellType == CellType.Numeric
                    && cell.ToString()?.Split('.').Length >= 3)
                    {
                        returnValue = ReturnStringDate(cell.DateCellValue);
                    }
                    else if (cell?.CellType == CellType.Formula)
                    {
                        if (cell?.CachedFormulaResultType == CellType.Numeric
                      && cell.NumericCellValue > 36526 &&
                      cell.NumericCellValue < 47484)
                        {
                            returnValue = ReturnStringDate(cell.DateCellValue);
                        }
                        else if (cell?.CachedFormulaResultType == CellType.Numeric)
                        {
                            returnValue = cell?.NumericCellValue.ToString();
                        }
                    }
                    else
                    {
                        returnValue = cell?.ToString();
                    }
                }
                return returnValue;
            }
            catch (Exception e)
            {
#if DEBUG
                Logger?.Error(e.Message);
                Logger?.Error(e.StackTrace);
#endif
                return default;
            }
        }

        public virtual void ReadValue()
        {
            throw new NotImplementedException("ReadValue()");
        }

        public static ICell GetFirstCellInMergedRegion(ICell cell)
        {
            if (cell != null && cell.IsMergedCell)
            {
                ISheet sheet = cell.Sheet;
                foreach (var region in sheet.MergedRegions)
                {
                    if (region.ContainsRow(cell.RowIndex) &&
                        region.ContainsColumn(cell.ColumnIndex))
                    {
                        IRow row = sheet.GetRow(region.FirstRow);
                        ICell firstCell = row?.GetCell(region.FirstColumn);
                        return firstCell;
                    }
                }
                return null;
            }
            return cell;
        }

        public virtual void InsertValue()
        {
            throw new NotImplementedException("InsertValue()");
        }

        public virtual void UpdateValue()
        {
            throw new NotImplementedException("UpdateValue()");
        }

        public virtual void DeleteValue()
        {
            if (ActiveSheet != null)
            {
                int countRows = LastViewedRow - FirstViewedRow + 1;
                for (int i = FirstViewedRow; i <= LastViewedRow; i++)
                {
                    var row = ActiveSheet.GetRow(i);
                    if (row != null)
                    {
                        var lastCol = row.LastCellNum;
                        if (lastCol < LastViewedColumn)
                        {
                            lastCol = (short)LastViewedColumn;
                        }
                        for (int ValueColumn = FirstViewedColumn; ValueColumn <= lastCol; ValueColumn++)
                        {
                            ICell cell = row.GetCell(ValueColumn);
                            if (ValueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                            {
                                row.RemoveCell(cell);
                            }
                            ProgressValue?.Report(ReturnProgress(i, countRows));
                        }
                    }
                }
            }
        }

        public static void SetCellValue(ISheet worksheet, int rowPosition,
            int columnPosition, string value)
        {
            IRow dataRow = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
            ICell cell = dataRow.GetCell(columnPosition) ?? dataRow.CreateCell(columnPosition);
            cell.SetCellValue(value);
        }

        public static void SetCellValue(ISheet worksheet, int rowPosition,
            int columnPosition, string value, CellType type)
        {
            IRow dataRow = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
            ICell cell = dataRow.GetCell(columnPosition) ?? dataRow.CreateCell(columnPosition, type);
            cell.SetCellValue(value);
        }
    }

    public class RowsView : ExchangeClass<IList<IRow>>
    {
        public string PathSource { get; set; }
        public override bool CloseStream => true;

        public RowsView(ExchangeOperation exchangeType, string activeSheetName, IList<IRow> exchangeValue,
            IProgress<int> progress = null) : base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
        }

        public override void ReadValue()
        {
            for (int i = FirstViewedRow; i <= LastViewedRow; i++)
            {
                ExchangeValue.Add(ActiveSheet.GetRow(i));
            }
        }

        public override void InsertValue()
        {
            UpdateValue(ActiveSheet.LastRowNum);
        }

        public override void UpdateValue()
        {
            UpdateValue(0);
        }


        public void UpdateValue(int StartRow = 0)
        {
            RowsView rowsView = new(ExchangeOperation.Read, this.ActiveSheetName, new List<IRow>(), null)
            {
                CloseStream = false
            };
            WrapperExcel tmpWrapper = new(PathSource, rowsView, Logger);
            tmpWrapper.Exchange();
            {
                rowsView.CloseStream = true;
                for (int i = rowsView.FirstViewedRow; i <= rowsView.LastViewedRow; i++)
                {
                    var row = rowsView.ActiveSheet.GetRow(i);
                    if (row != null)
                    {
                        ChangedNPOI.ChangedCopyRow(rowsView.ActiveSheet, i, this.ActiveSheet, StartRow + i);
                    }
                }
            }
        }
    }

    public class MatrixView : ExchangeClass<IList<string[]>>
    {

        public MatrixView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string[]> exchangeValue, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
        }

        private void AddValue()
        {
            if (ExchangeValue != null)
            {
                int rowsCount = ActiveSheet.RowsCount();
                for (int i = 0; i < ExchangeValue.Count; i++)
                {
                    IRow row = ActiveSheet.CreateRow(i + FirstViewedRow + rowsCount);
                    for (int j = 0; j < ExchangeValue[i].Length; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j + FirstViewedColumn);
                        cell.SetCellValue(ExchangeValue[i][j]);
                    }
                    ProgressValue?.Report(ReturnProgress(i, ExchangeValue.Count));
                }
            }
        }

        public override void InsertValue()
        {
            AddValue();
        }

        public override void UpdateValue()
        {
            int fRow = FirstViewedRow;
            int lRow = LastViewedRow;
            int total = lRow - fRow;
            for (int i = lRow; i >= fRow; i--)
            {
                var row = ActiveSheet.GetRow(i);
                if (row != null)
                {
                    ActiveSheet.RemoveRow(row);
                }
                ProgressValue?.Report(ReturnProgress(i - fRow, total));
            }
            AddValue();
        }

        private string[] GetStringFromRow(int i, int firstViewedColumn, int lastViewedColumn)
        {
            var row = ActiveSheet.GetRow(i);
            List<string> tmp = new();
            if (row != null)
            {
                var lastCol = row.LastCellNum;
                if (lastCol < lastViewedColumn)
                {
                    lastCol = (short)lastViewedColumn;
                }
                for (int valueColumn = firstViewedColumn; valueColumn <= lastCol; valueColumn++)
                {
                    ICell cell;
                    cell = row.GetCell(valueColumn);
                    if (valueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                    {
                        tmp.Add(GetCellValue(cell));
                    }
                }
            }
            return tmp.ToArray();
        }


        public override void ReadValue()
        {
            if (FirstViewedRow == 0 &&
                LastViewedRow == 0 &&
                LastViewedColumn == 0 &&
                FirstViewedColumn == 0)
            {
                ReadValueHoleSheet();
            }
            else
            {
                ReadValueWithBorders();
            }
        }

        private void ReadValueHoleSheet() //Fast
        {
            ExchangeValue = new List<string[]>();
            int firstViewedRow = FirstViewedRow;
            //int lastViewedRow = LastViewedRow;
            int lastViewedColumn = LastViewedColumn;
            int firstViewedColumn = FirstViewedColumn;
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
                ExchangeValue = tmpListString;
            }
        }

        private void ReadValueWithBorders() //Slow
        {
            ExchangeValue = new List<string[]>();
            int lastViewedRow = LastViewedRow;
            int lastViewedColumn = LastViewedColumn;
            int firstViewedColumn = FirstViewedColumn;
            int firstViewedRow = FirstViewedRow;
            List<string[]> tmp = new();
            if (ActiveSheet != null)
            {
                int countValue = lastViewedRow - firstViewedRow + 1;
                for (int i = firstViewedRow; i <= lastViewedRow; i++)
                {
                    tmp.Add(GetStringFromRow(i, firstViewedColumn, lastViewedColumn));
                    ProgressValue?.Report(ReturnProgress(i, countValue));
                }
                ExchangeValue = tmp;
            }
        }
    }

    public class ListView : ExchangeClass<IList<string>>
    {
        private readonly MatrixView matrix;
        public ListView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string> exchangeValue, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, progress)
        {
            matrix = new MatrixView(exchangeType, activeSheetName,
            Extension.ConvnertToMatrix(exchangeValue), progress)
            {
                FirstViewedRow = this.FirstViewedRow,
                FirstViewedColumn = this.FirstViewedColumn,
                LastViewedColumn = this.LastViewedColumn,
                LastViewedRow = this.LastViewedRow
            };
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.FirstViewedRow = this.FirstViewedRow;
            matrix.FirstViewedColumn = this.FirstViewedColumn;
            matrix.LastViewedColumn = this.LastViewedColumn;
            matrix.LastViewedRow = this.LastViewedRow;
            matrix.InsertValue();
        }

        public override void ReadValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.FirstViewedRow = this.FirstViewedRow;
            matrix.FirstViewedColumn = this.FirstViewedColumn;
            matrix.LastViewedColumn = this.LastViewedColumn;
            matrix.LastViewedRow = this.LastViewedRow;
            matrix.ReadValue();
            ExchangeValue = Extension.ConvertToList(matrix.ExchangeValue);
        }

        public override void UpdateValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.FirstViewedRow = this.FirstViewedRow;
            matrix.FirstViewedColumn = this.FirstViewedColumn;
            matrix.LastViewedColumn = this.LastViewedColumn;
            matrix.LastViewedRow = this.LastViewedRow;
            matrix.UpdateValue();
        }
    }

    public class DictionaryView : ExchangeClass<IDictionary<string, string[]>>
    {
        private readonly MatrixView matrix;
        public DictionaryView(ExchangeOperation exchangeType, string activeSheetName,
            IDictionary<string, string[]> exchangeValue, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, progress)
        {
            matrix = new MatrixView(exchangeType, activeSheetName,
            Extension.ConvnertToMatrix(exchangeValue), progress);
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.InsertValue();
        }

        public override void ReadValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.ReadValue();
            ExchangeValue = Extension.ConvertToDictionary(matrix.ExchangeValue);
        }

        public override void UpdateValue()
        {
            matrix.ActiveSheet = this.ActiveSheet;
            matrix.UpdateValue();
        }

    }

    

        

        
    


    public static class ExcelExchange
    {
        // статический класс для записи и чтения данных одной строкой
        /// <summary>
        /// The TaskAddToExcel.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string}"/>.</param>
        public static void TaskAddToExcel(string pathToFile, string sheetName, List<string> values)
        {
            Task AddValueToExcel = Task.Run(() =>
            {
                ListView listView = new(ExchangeOperation.Insert, sheetName, values, null)
                {
                    ExchangeValue = values
                };
                WrapperExcel wrapper = new(pathToFile, listView, null) { };
                wrapper.Exchange();
            });
        }

        /// <summary>
        /// The TaskAddToExcel.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        public static void TaskAddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            try
            {
                Task AddValueToExcel = Task.Run(() =>
                  {
                      AddToExcel(pathToFile, sheetName, values);
                  });
            }
            catch (Exception e)
            {
#if DEBUG
                Debug.WriteLine(e.Message);
#endif
            }
        }

        public static async Task TaskAddToExcelAsync(string pathToFile, string sheetName, List<string[]> values)
        {
            await Task.Run(() =>
            {
                AddToExcel(pathToFile, sheetName, values);
            });
        }

        /// <summary>
        /// The AddToExcel
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        public static void AddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            MatrixView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            WrapperExcel wrapper = new(pathToFile, listView, null)
            {
                //ActiveSheetName = sheetName,
                //exchangeClass = listView
            };
            wrapper.Exchange();
        }

        /// <summary>
        /// The AddToExcel.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string}"/>.</param>
        public static void AddToExcel(string pathToFile, string sheetName, List<string> values)
        {
            ListView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            WrapperExcel wrapper = new(pathToFile, listView, null)
            {
                //ActiveSheetName = sheetName,
                //exchangeClass = listView
            };
            wrapper.Exchange();
        }

        /// <summary>
        /// The GetFromExcel.
        /// </summary>
        /// <typeparam name="ReturnType">.</typeparam>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <returns>The <see cref="ReturnType"/>.</returns>
        public static ReturnType GetFromExcel<ReturnType>(string pathToFile, string sheetName, int firstRow = 0, int firstCol = 0) where ReturnType : new()
        {
            ReturnType returnValue = new();
            if (returnValue is List<string> rL)
            {
                var exchangeClass = new ListView(ExchangeOperation.Read, sheetName, rL, null)
                {
                    FirstViewedColumn = firstCol,
                    FirstViewedRow = firstRow
                };
                WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                var exchangeClass = new DictionaryView(ExchangeOperation.Read, sheetName, rD, null)
                {
                    FirstViewedColumn = firstCol,
                    FirstViewedRow = firstRow
                };
                WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else if (returnValue is List<string[]> rM)
            {
                var exchangeClass = new MatrixView(ExchangeOperation.Read, sheetName, rM, null)
                {
                    FirstViewedColumn = firstCol,
                    FirstViewedRow = firstRow
                };
                WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else
            {
                throw new TypeUnloadedException("No handler for type");
            }
        }
    }


}