#define DEBUG

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

using NPOI.POIFS.Crypt;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
namespace WrapperNetPOI.Excel
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class IsExternalInit
    { }
}
namespace WrapperNetPOI.Excel
{
    public class Border // in developing
    {
        public ISheet ActiveSheet { get; set; }
        private int? firstColumn;
        private int? firstRow;
        private int? lastRow;
        private int? lastColumn;
        public int FirstRow
        {
            set
            {
                firstRow = value;
            }
            get
            {
                if (firstRow == null)
                {
                    return (ActiveSheet?.FirstRowNum) ?? 0;
                }
                else
                {
                    return firstRow ?? 0;
                }
            }
        }
        public int FirstColumn
        {
            set
            {
                firstColumn = value;
            }
            get
            {
                if (firstColumn == null)
                {
                    firstColumn = ActiveSheet?.GetRow(ActiveSheet?.FirstRowNum ?? 0)?.FirstCellNum;
                    if (firstColumn == -1)
                    {
                        firstColumn = 0;
                        return 0;
                    }
                    else
                    {
                        return firstColumn ?? 0;
                    }
                }
                else
                {
                    return firstColumn ?? 0;
                }
            }
        }
        public int LastRow
        {
            set
            {
                lastRow = value;
            }
            get
            {
                if (lastRow == null || lastRow == 0)
                {
                    return (ActiveSheet?.LastRowNum) ?? 0;
                }
                else
                {
                    return lastRow ?? 0;
                }
            }
        }
        public int LastColumn
        {
            set
            {
                lastColumn = value;
            }
            get
            {
                if (lastColumn == null)
                {
                    lastColumn = ActiveSheet?.GetRow(ActiveSheet?.FirstRowNum ?? 0)?.LastCellNum;
                    if (lastColumn == -1)
                    {
                        lastColumn = 0;
                        return 0;
                    }
                    else
                    {
                        return lastColumn ?? 0;
                    }
                }
                else
                {
                    return lastColumn ?? 0;
                }
            }
        }

        public Border()
        { }

        public Border(int? firstRow = null, int? firstColumn = null, int? lastRow = null, int? lastColumn = null)
        {
            CorrectBorder(firstRow, firstColumn, lastRow, lastColumn);
        }

        public void CorrectBorder(int? firstRow = null, int? firstColumn = null, int? lastRow = null, int? lastColumn = null)
        {
            if (firstRow != null)
            {
                FirstRow = firstRow ?? (int)firstRow;
            }
            if (firstColumn != null)
            {
                FirstColumn = firstColumn ?? (int)firstColumn;
            }
            if (lastColumn != null)
            {
                LastColumn = lastColumn ?? (int)lastColumn;
            }
            if (lastRow != null)
            {
                LastRow = lastRow ?? (int)lastRow;
            }
        }

        public int Row(int i)
        {
            if (firstRow == null && lastRow == null)
            {
                return i;
            }
            else if (firstRow != null && lastRow == null)
            {
                return i + FirstRow;
            }
            else if (firstRow != null && lastRow != null)
            {
                if (i + FirstRow <= LastRow)
                {
                    return i + FirstRow;
                }
                else
                {
                    return LastRow;
                }
            }
            else
            {
                return FirstRow;
            }
        }

        public int Column(int i)
        {
            if (firstColumn == null && lastColumn == null)
            {
                return i;
            }
            else if (firstColumn != null && lastColumn == null)
            {
                return i + FirstColumn;
            }
            else if (firstColumn != null && lastColumn != null)
            {
                if (i + FirstColumn <= LastColumn)
                {
                    return i + FirstColumn;
                }
                else
                {
                    return LastColumn;
                }
            }
            else
            {
                return FirstColumn;
            }
        }
    }
    internal class GetCellValue
    {
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
        public static IList<string[]> ConvertToMatrix(IList<string> list)
        {
            return list?.Select(x => new string[] { x }).ToList();
        }
        public static IList<string[]> ConvertToMatrix(IDictionary<string, string[]> dict)
        {
            return dict?.SelectMany(x => x.Value.ToList(), (key, value) => new string[] { key.Key, value }).ToList();
        }
        public static IList<string> ConvertToList(IList<string[]> list)
        {
            return list.Select(x => x.FirstOrDefault()).ToList();
        }
        public static IDictionary<string, string[]> ConvertToDictionary(IList<string[]> list)
        {
            Dictionary<string, string[]> dict = new();
            var groupValue = list?.Where(x => x.Length >= 2).Where(x => x[0] != null).GroupBy(y => y[0], (value) => value[1]);
            foreach (var group in groupValue)
            {
                dict[group.Key] = group.Select(x => x).ToArray();
            }
            return dict;
        }
    }

    public static class MatrixConvert<T>
    {

        public static IList<T[]> ConvertToMatrix(IList<T> list)
        {
            return list?.Select(x => new T[] { x }).ToList();
        }
        public static IList<T[]> ConvertToMatrix(IDictionary<T, T[]> dict)
        {
            return dict?.SelectMany(x => x.Value.ToList(), (key, value) => new T[] { key.Key, value }).ToList();
        }
        public static IList<T> ConvertToList(IList<T[]> list)
        {
            return list.Select(x => x.FirstOrDefault()).ToList();
        }
        public static IDictionary<T, T[]> ConvertToDictionary(IList<T[]> list)
        {
            Dictionary<T, T[]> dict = new();
            var groupValue = list?.Where(x => x.Length >= 2).Where(x => x[0] != null).GroupBy(y => y[0], (value) => value[1]);
            foreach (var group in groupValue)
            {
                dict[group.Key] = group.Select(x => x).ToArray();
            }
            return dict;
        }
    }

    public interface IExchangeExcel : IExchange
    {
        string ActiveSheetName { set; get; }
        //int FirstViewedRow { get; }
        //int LastViewedRow { get; }
        //int FirstViewedColumn { get; }
        //int LastViewedColumn { get; }
        ISheet ActiveSheet { set; get; }
        IWorkbook Workbook { set; get; }
    }
    public abstract class NewBaseType
    {
        public static int ReturnProgress(int number, int total)
        {
            if (total != 0)
            {
                return number * 100 / total;
            }
            else
            {
                return 0;
            }
        }
    }
    public abstract class ExchangeClass<Tout> : NewBaseType, IExchangeExcel
    {
        public Border WorkbookBorder { set; get; }
        public IWorkbook Workbook { set; get; }
        private string Password { set; get; }
        public virtual bool CloseStream { set; get; } = true;
        public IProgress<int> ProgressValue { set; get; }
        public ILogger Logger { set; get; }
        public string ActiveSheetName { set; get; }
        public ExchangeOperation ExchangeOperationEnum { set; get; }
        public string[] SheetsNames { set; get; }
        private ISheet activeSheet;

        public virtual ISheet ActiveSheet
        {
            set
            {
                activeSheet = value;
                WorkbookBorder ??= new();
                WorkbookBorder.ActiveSheet = activeSheet;
            }
            get
            {
                return activeSheet;
            }
        }
        //public int FirstViewedRow => WorkbookBorder.FirstRow;
        //public int FirstViewedColumn => WorkbookBorder.FirstColumn;
        //public int LastViewedRow => WorkbookBorder.LastRow;
        //public int LastViewedColumn => WorkbookBorder.LastColumn;
        public Tout ExchangeValue { set; get; }
        public Action ExchangeValueFunc { set; get; }
        protected ExchangeClass(ExchangeOperation exchange, string activeSheetName, Border border = null, IProgress<int> progress = null)
        {
            WorkbookBorder = border;
            ExchangeOperationEnum = exchange;
            ActiveSheetName = activeSheetName;
            ProgressValue = progress;
        }
        public virtual void GetInternallyObject(Stream tmpStream, bool addNewWorkbook)
        {
            FileStream fs = default;
            if (Password
                != null)
            {
                NPOI.POIFS.FileSystem.POIFSFileSystem nfs =
                new(fs);
                EncryptionInfo info = new(nfs);
                Decryptor dc = Decryptor.GetInstance(info);
                //bool b = dc.VerifyPassword(Password);
                dc.VerifyPassword(Password);
                tmpStream = dc.GetDataStream(nfs);
            }
            if (addNewWorkbook)
            {
                if (ActiveSheetName==null)
                {
                    throw new ArgumentNullException("The sheet name cannot be null");
                }
                Workbook = new XSSFWorkbook();
                Workbook.CreateSheet(ActiveSheetName);
                ActiveSheet = Workbook.GetSheet(ActiveSheetName);
            }
            else
            {
                Workbook = WorkbookFactory.Create(tmpStream);
                int SheetsCount = Workbook.NumberOfSheets;
                bool getValue = false;
                for (int i = 0; i < Workbook.NumberOfSheets; i++)
                {
                    if (Workbook.GetSheetAt(i).SheetName == ActiveSheetName)
                    {
                        if (Workbook.GetSheet(ActiveSheetName) is ISheet activeSheet)
                        {
                            ActiveSheet = activeSheet;
                            getValue = true;
                        }
                        break;
                    }
                }
                if (!getValue && SheetsCount != 0)
                {
                    //throw (new ArgumentOutOfRangeException($"No page found with that name-{ActiveSheetName}"));
                    ActiveSheet = Workbook.GetSheetAt(0);
                    // search first if not found
                }
            }
            //exchangeClass.ActiveSheet = ActiveSheet;
            SheetsNames = ReturnSheetsNames();
            ExchangeValueFunc();
        }

        public string[] ReturnSheetsNames()
        {
            List<string> tmp = new();
            for (int i = 0; i < Workbook.NumberOfSheets; i++)
            {
                tmp.Add(Workbook.GetSheetAt(i).SheetName);
            }
            return tmp.ToArray();
        }

        /// <summary>
        /// $Return Date by dd.mm.yyyy$
        /// </summary>
        /// <param name="date">The date<see cref="DateTime"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string ReturnStringDate(DateTime date)
        {
            var day = string.Format("{0:D2}", date.Day);
            var mounth = string.Format("{0:D2}", date.Month);
            var year = string.Format("{0:D4}", date.Year);
            return $"{day}.{mounth}.{year}";
        }
        public virtual string GetCellValue(ICell cell)
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
                    returnValue = new WrapperCell(cell).GetValue<string>();
                }
                return returnValue;
            }
            catch (Exception e)
            {
                //#if DEBUG
                Logger?.Error(e.Message);
                Logger?.Error(e.StackTrace);
                //#endif
                return default;
            }
        }
        public virtual void ReadValue()
        {
            throw new NotImplementedException("ReadValue()");
        }
        public static ICell GetFirstCellInMergedRegion(ICell cell)
        {
            if (cell?.IsMergedCell == true)
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
                int countRows = WorkbookBorder.LastRow - WorkbookBorder.FirstRow + 1;
                for (int i = WorkbookBorder.FirstRow; i <= WorkbookBorder.LastRow; i++)
                {
                    var row = ActiveSheet.GetRow(i);
                    if (row != null)
                    {
                        var lastCol = row.LastCellNum;
                        if (lastCol < WorkbookBorder.LastColumn)
                        {
                            lastCol = (short)WorkbookBorder.LastColumn;
                        }
                        for (int ValueColumn = WorkbookBorder.FirstColumn; ValueColumn <= lastCol; ValueColumn++)
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
        public RowsView(ExchangeOperation exchangeType, string activeSheetName = "", IList<IRow> exchangeValue = null,
            IProgress<int> progress = null) : base(exchangeType, activeSheetName, new Border(), progress)
        {
            ExchangeValue = exchangeValue ?? new List<IRow>();
        }
        public override void ReadValue()
        {
            for (int i = WorkbookBorder.FirstRow; i <= WorkbookBorder.LastRow; i++)
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
            RowsView rowsView = new(ExchangeOperation.Read, ActiveSheetName, new List<IRow>(), null)
            {
                CloseStream = true
                //CloseStream = false
            };
            WrapperExcel tmpWrapper = new(PathSource, rowsView, Logger);
            tmpWrapper.Exchange();
            {
                rowsView.CloseStream = true;
                for (int i = rowsView.WorkbookBorder.FirstRow; i <= rowsView.WorkbookBorder.LastRow; i++)
                {
                    var row = rowsView.ActiveSheet.GetRow(i);
                    if (row != null)
                    {
                        ChangedNPOI.ChangedCopyRow(rowsView.ActiveSheet, i, ActiveSheet, StartRow + i);
                    }
                }
            }
        }
    }
    public class MatrixViewGeneric<T> : ExchangeClass<IList<T[]>>
    {
        public MatrixViewGeneric(ExchangeOperation exchangeType, string activeSheetName,
            IList<T[]> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
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
                    IRow row = ActiveSheet.CreateRow(i + WorkbookBorder.FirstRow + rowsCount);
                    for (int j = 0; j < ExchangeValue[i].Length; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j + WorkbookBorder.FirstColumn);
                        new WrapperCell(cell).SetValue(ExchangeValue[i][j]);
                        //cell.SetCellValue(ExchangeValue[i][j]);
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
            int fRow = WorkbookBorder.FirstRow;
            int lRow = WorkbookBorder.LastRow;
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
            List<string> tmp = new ();
            if (row != null)
            {
                var lastCol = row.LastCellNum;
                if (lastCol < lastViewedColumn)
                {
                    lastCol = (short)lastViewedColumn;
                }
                for (int valueColumn = firstViewedColumn; valueColumn <= lastCol; valueColumn++)
                {
                    ICell cell = row.GetCell(valueColumn);
                    if (valueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                    {
                        tmp.Add(GetCellValue(cell));
                    }
                }
            }
            return tmp.ToArray();
        }
        private T[] GetArrayFromRow(int i, int firstViewedColumn, int lastViewedColumn)
        {
            var row = ActiveSheet.GetRow(i);
            List<T> tmp = new();
            if (row != null)
            {
                var lastCol = row.LastCellNum;
                if (lastCol < lastViewedColumn)
                {
                    lastCol = (short)lastViewedColumn;
                }
                for (int valueColumn = firstViewedColumn; valueColumn <= lastCol; valueColumn++)
                {
                    ICell cell = row.GetCell(valueColumn);
                    if (valueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                    {
                        //WrapperCell wrapper = new(cell);
                        //ConvertType convert = new();
                        new WrapperCell(cell).GetValue<T>(out var value);
                        tmp.Add(value);
                    }
                }
            }
            return tmp.ToArray();
        }
        public override void ReadValue()
        {
            if (WorkbookBorder.FirstRow == 0 &&
                WorkbookBorder.LastRow == 0 &&
                WorkbookBorder.LastColumn == 0 &&
                WorkbookBorder.FirstColumn == 0)
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
            ExchangeValue = new List<T[]>();
            int firstViewedRow = WorkbookBorder.FirstRow;
            //int lastViewedRow = LastViewedRow;
            int lastViewedColumn = WorkbookBorder.LastColumn;
            int firstViewedColumn = WorkbookBorder.FirstColumn;
            List<T[]> tmpListString = new();
            if (ActiveSheet != null)
            {
                int i = 0;
                foreach (IRow value in ActiveSheet)
                {
                    IRow row = value;
                    if (row.RowNum > i)
                    {
                        do
                        {
                            tmpListString.Add(Array.Empty<T>());
                            i++;
                        }
                        while (row.RowNum != i);
                    }
                    //ConvertType convert = new();
                    tmpListString.Add(row.Select(x => new WrapperCell(x).GetValue<T>()).ToArray());
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
            ExchangeValue = new List<T[]>();
            int lastViewedRow = WorkbookBorder.LastRow;
            int lastViewedColumn = WorkbookBorder.LastColumn;
            int firstViewedColumn = WorkbookBorder.FirstColumn;
            int firstViewedRow = WorkbookBorder.FirstRow;
            List<T[]> tmp = new();
            if (ActiveSheet != null)
            {
                int countValue = lastViewedRow - firstViewedRow + 1;
                for (int i = firstViewedRow; i <= lastViewedRow; i++)
                {
                    tmp.Add(GetArrayFromRow(i, firstViewedColumn, lastViewedColumn));
                    ProgressValue?.Report(ReturnProgress(i, countValue));
                }
                ExchangeValue = tmp;
            }
        }
    }

    /*
    public class MatrixView : ExchangeClass<IList<string[]>>
    {
        public MatrixView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string[]> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
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
                    IRow row = ActiveSheet.CreateRow(i + WorkbookBorder.FirstRow + rowsCount);
                    for (int j = 0; j < ExchangeValue[i].Length; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j + WorkbookBorder.FirstColumn);
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
            int fRow = WorkbookBorder.FirstRow;
            int lRow = WorkbookBorder.LastRow;
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
                    ICell cell = row.GetCell(valueColumn);
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
            if (WorkbookBorder.FirstRow == 0 &&
                WorkbookBorder.LastRow == 0 &&
                WorkbookBorder.LastColumn == 0 &&
                WorkbookBorder.FirstColumn == 0)
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
            int firstViewedRow = WorkbookBorder.FirstRow;
            //int lastViewedRow = LastViewedRow;
            int lastViewedColumn = WorkbookBorder.LastColumn;
            int firstViewedColumn = WorkbookBorder.FirstColumn;
            List<string[]> tmpListString = new();
            if (ActiveSheet != null)
            {
                int i = 0;
                foreach (IRow value in ActiveSheet)
                {
                    IRow row = value;
                    if (row.RowNum > i)
                    {
                        do
                        {
                            tmpListString.Add(Array.Empty<string>());
                            i++;
                        }
                        while (row.RowNum != i);
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
            int lastViewedRow = WorkbookBorder.LastRow;
            int lastViewedColumn = WorkbookBorder.LastColumn;
            int firstViewedColumn = WorkbookBorder.FirstColumn;
            int firstViewedRow = WorkbookBorder.FirstRow;
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
    
    */
    /*
    public class ListView : ExchangeClass<IList<string>>
    {
        private readonly MatrixViewGeneric<string> matrix;
        public ListView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
        {
            matrix = new MatrixViewGeneric<string>(exchangeType, activeSheetName,
            Extension.ConvertToMatrix(exchangeValue), border, progress);
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.InsertValue();
        }
        public override void ReadValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.ReadValue();
            ExchangeValue = Extension.ConvertToList(matrix.ExchangeValue);
        }
        public override void UpdateValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.UpdateValue();
        }
    }

    */

    public class ListViewGeneric<T> : ExchangeClass<IList<T>>
    {
        private readonly MatrixViewGeneric<T> matrix;
        public ListViewGeneric(ExchangeOperation exchangeType, string activeSheetName,
            IList<T> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
        {
            matrix = new MatrixViewGeneric<T>(exchangeType, activeSheetName,
            MatrixConvert<T>.ConvertToMatrix(exchangeValue), border, progress);
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.InsertValue();
        }
        public override void ReadValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.ReadValue();
            ExchangeValue = MatrixConvert<T>.ConvertToList(matrix.ExchangeValue);
        }
        public override void UpdateValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.WorkbookBorder = WorkbookBorder;
            matrix.UpdateValue();
        }
    }


    public class DictionaryView : ExchangeClass<IDictionary<string, string[]>>
    {
        private readonly MatrixViewGeneric<string> matrix;
        public DictionaryView(ExchangeOperation exchangeType, string activeSheetName,
            IDictionary<string, string[]> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
        {
            matrix = new MatrixViewGeneric<string>(exchangeType, activeSheetName,
            Extension.ConvertToMatrix(exchangeValue), border, progress);
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.InsertValue();
        }
        public override void ReadValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.ReadValue();
            ExchangeValue = Extension.ConvertToDictionary(matrix.ExchangeValue);
        }
        public override void UpdateValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.UpdateValue();
        }
    }

    public class DictionaryViewGeneric<T> : ExchangeClass<IDictionary<T, T[]>>
    {
        private readonly MatrixViewGeneric<T> matrix;
        public DictionaryViewGeneric(ExchangeOperation exchangeType, string activeSheetName,
            IDictionary<T, T[]> exchangeValue, Border border = null, IProgress<int> progress = null) :
            base(exchangeType, activeSheetName, border, progress)
        {
            matrix = new MatrixViewGeneric<T>(exchangeType, activeSheetName,
            MatrixConvert<T>.ConvertToMatrix(exchangeValue), border, progress);
        }
        public override void InsertValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.InsertValue();
        }
        public override void ReadValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.ReadValue();
            ExchangeValue = MatrixConvert<T>.ConvertToDictionary(matrix.ExchangeValue);
        }
        public override void UpdateValue()
        {
            matrix.ActiveSheet = ActiveSheet;
            matrix.UpdateValue();
        }
    }
}