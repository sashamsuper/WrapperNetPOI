using NPOI.HSSF.UserModel;
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
using Microsoft.CSharp;
using System.Text.Json;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using MathNet.Numerics.Optimization;

namespace WrapperNetPOI
{

    public enum ExchangeOperation
    {
        Insert,
        Read,
        Update,
        Delete
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

        public static IDictionary<string,string[]> ConvertToDictionary(IList<string[]> list)
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
        IProgress<double> ProgressValue { set; get; }
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

    /// <summary>
    /// Defines the <see cref="ExchangeClass" />.
    /// </summary>
    public abstract class ExchangeClass<Tout> : IExchange
    {
        public virtual bool CloseStream { set; get; } = true;

        public IProgress<double> ProgressValue { set; get; }

        public ILogger Logger { set; get; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeClass"/> class.
        /// </summary>
        public ExchangeClass(ExchangeOperation exchange, string activeSheetName, IProgress<double> progress)
        {
            ExchangeOperationEnum = exchange;
            ActiveSheetName = activeSheetName;
            ProgressValue = progress;
        }

        public string ActiveSheetName { set; get; }
        public static double ReturnProgress(int i, int total, int firstValue = 0)
        {
            if (total != 0)
            {
                return (i - firstValue + 1) / ((double)(total)) * 100.0;
            }
            else
            {
                return 0;
            }
        }

        public ExchangeOperation ExchangeOperationEnum { set; get; }
        /// <summary>
        /// ActiveSheet
        /// </summary>
        public ISheet ActiveSheet { set; get; }

        // <summary>
        /// The initial line with which data is entered / viewed
        /// </summary>
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
                    
                    return (ActiveSheet!=null)?ActiveSheet.FirstRowNum:0;
                }
                else 
                {
                    return firstViewedRow??0;
                }
            }
        }
        private int? firstViewedRow;

        /// <summary>
        /// The initial column from which data is entered/viewed
        /// </summary>
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
        private int ? firstViewedColumn;
        
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
                    return (ActiveSheet != null) ? ActiveSheet.LastRowNum:0;
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

        /// <summary>
        /// $Column in which data is entered/viewed$
        /// </summary>
        //public int ValueColumn { set; get; } = 1;

        /// <summary>
        /// Gets or sets the ExchangeValue.
        /// </summary>
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
            return day + "." + mounth + "." + year;
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
                Logger?.Error(e.Message);
                Logger?.Error(e.StackTrace);
                return default;
            }
        }

        /// <summary>
        /// The FindValue.
        /// </summary>
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


        /// <summary>
        /// The AddValue.
        /// </summary>
        public virtual void InsertValue()
        {
            throw new NotImplementedException("InsertValue()");
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        public virtual void UpdateValue()
        {
            throw new NotImplementedException("UpdateValue()");
        }

        public virtual void DeleteValue()
        {
            if (ActiveSheet != null)
            {
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
                            ProgressValue?.Report((i - FirstViewedRow) / (LastViewedRow - FirstViewedRow) * 100.0);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// The SetCellValue.
        /// </summary>
        /// <param name="worksheet">The worksheet<see cref="ISheet"/>.</param>
        /// <param name="rowPosition">The rowPosition<see cref="int"/>.</param>
        /// <param name="columnPosition">The columnPosition<see cref="int"/>.</param>
        /// <param name="value">The value<see cref="string"/>.</param>
        public static void SetCellValue(ISheet worksheet, int rowPosition, int columnPosition, string value)
        {
            IRow dataRow = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
            ICell cell = dataRow.GetCell(columnPosition) ?? dataRow.CreateCell(columnPosition);
            cell.SetCellValue(value);
        }

        /// <summary>
        /// The SetCellValue.
        /// </summary>
        /// <param name="worksheet">The worksheet<see cref="ISheet"/>.</param>
        /// <param name="rowPosition">The rowPosition<see cref="int"/>.</param>
        /// <param name="columnPosition">The columnPosition<see cref="int"/>.</param>
        /// <param name="value">The value<see cref="string"/>.</param>
        /// <param name="type">The type<see cref="CellType"/>.</param>
        public static void SetCellValue(ISheet worksheet, int rowPosition, int columnPosition, string value, CellType type)
        {
            IRow dataRow = worksheet.GetRow(rowPosition) ?? worksheet.CreateRow(rowPosition);
            ICell cell = dataRow.GetCell(columnPosition) ?? dataRow.CreateCell(columnPosition, type);
            cell.SetCellValue(value);
        }
    }


    public class RowsView : ExchangeClass<IList<IRow>>
    {
        public int CountRows { get; set; }
        public string PathSource { get; set; }
        public override bool CloseStream => true;

        public RowsView(ExchangeOperation exchangeType, string activeSheetName, IList<IRow> exchangeValue,
            IProgress<double> progress) : base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
            //ValueColumn = 0;
        }
        public override void ReadValue()
        {
            for (int i = FirstViewedRow; i < CountRows; i++)
            {
                ExchangeValue.Add(ActiveSheet.GetRow(i));
            }

        }
        public override void InsertValue()
        {
            UpdateValue(this.CountRows);
        }

        public override void UpdateValue()
        {
            UpdateValue(0);
        }


        public void UpdateValue(int StartRow = 0)
        {
            RowsView rowsView = new(ExchangeOperation.Read, this.ActiveSheetName, new List<IRow>(), null)
            {
                CountRows = this.CountRows,
                CloseStream = false
            };
            Wrapper tmpWrapper = new(PathSource, rowsView, Logger);
            tmpWrapper.Exchange();
            {
                rowsView.CloseStream = true;
                for (int i = 0; i < CountRows; i++)
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

    /// <summary>
    /// Defines the <see cref="MatrixView" />.
    /// </summary>
    public class MatrixView : ExchangeClass<IList<string[]>>
    {
        // занесение спика массива=строкам
        /// <summary>
        /// Initializes a new instance of the <see cref="MatrixView"/> class.
        /// </summary>
        public MatrixView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string[]> exchangeValue, IProgress<double> progress) :
            base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
            //ValueColumn = 0;
        }

        /// <summary>
        /// The AddValue.
        /// </summary>
        /// <param name="startRow">The startRow<see cref="int"/>.</param>
        private void AddValue()
        {
            if (ExchangeValue != null)
            {
                int rowsCount = ActiveSheet.RowsCount();
                for (int i = 0; i < ExchangeValue.Count; i++)
                {
                    IRow row = ActiveSheet.CreateRow(i + FirstViewedRow+rowsCount);
                    for (int j = 0; j < ExchangeValue[i].Length; j++)
                    {
                        ICell cell = row.GetCell(j) ?? row.CreateCell(j+FirstViewedColumn);
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

        /// <summary>
        /// The AddValue.
        /// </summary>
        public override void UpdateValue()
        {
            int fRow = FirstViewedRow;
            int lRow = LastViewedRow;
            for (int i = lRow; i >= fRow; i--)
            {
                var row = ActiveSheet.GetRow(i);
                if (row != null)
                {
                    ActiveSheet.RemoveRow(row);
                }
            }
            AddValue();
        }


        public string[] GetStringFromRow(int i,int firstViewedColumn, int lastViewedColumn)
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
                        cell =row.GetCell(valueColumn);
                        if (valueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                        {
                            tmp.Add(GetCellValue(cell));
                        }
                    }
                }
                //ProgressValue?.Report(ReturnProgress(i, lastViewedRow - firstViewedRow + 1));
                //exchangeValue.Add(tmpListString.ToArray());
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




        public void ReadValueHoleSheet() //Fast
        {
            ExchangeValue = new List<string[]>();
            int firstViewedRow = FirstViewedRow;
            //int lastViewedRow = LastViewedRow;
            int lastViewedColumn = LastViewedColumn;
            int firstViewedColumn = FirstViewedColumn;
            List<string[]> tmpListString = new();

            if (ActiveSheet != null)
            {
                Stopwatch stopwatch = new();
                Debug.WriteLine("stopwatch");
                stopwatch.Start();
                int i = 0;
                foreach (IRow value in ActiveSheet)
                {

                    IRow row=value;
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
                    if (i % 1000 == 0)
                    {
                        Debug.WriteLine(i);
                    }
                }
                ExchangeValue = tmpListString;
                stopwatch.Stop();
                Debug.WriteLine($"StopWatch millisecondes-{stopwatch.ElapsedMilliseconds}");

            }
        }

    /// <summary>
    /// The GetValue.
    /// </summary>
    public void ReadValueWithBorders() //Slow
        {
            ExchangeValue = new List<string[]>();
            int lastViewedRow = LastViewedRow;
            int lastViewedColumn = LastViewedColumn;
            int firstViewedColumn = FirstViewedColumn;
            int firstViewedRow = FirstViewedRow;
            List<string[]> tmp=new();
            if (ActiveSheet != null)
            {
                Stopwatch stopwatch = new();
                Debug.WriteLine("stopwatch");
                stopwatch.Start();
                for (int i = firstViewedRow; i <= lastViewedRow; i++)
                {
                    tmp.Add(GetStringFromRow(i, firstViewedColumn, lastViewedColumn));
                }
                ExchangeValue=tmp;
                stopwatch.Stop();
                Debug.WriteLine($"StopWatch millisecondes-{stopwatch.ElapsedMilliseconds}");
            }
        }
    }

    public class ListView : ExchangeClass<IList<string>>
    {
        private readonly MatrixView matrix;
        public ListView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string> exchangeValue, IProgress<double> progress) :
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


    /*
    /// <summary>
    /// Defines the <see cref="ListView" />.
    /// </summary>
    public class LListView : ExchangeClass<IList<string>>
    {
        // обновление по листу
        /// <summary>
        /// Initializes a new instance of the <see cref="ListView"/> class.
        /// </summary>
        public LListView(ExchangeOperation exchangeType, string activeSheetName,
            IList<string> exchangeValue, IProgress<double> progress) :
            base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
            ValueColumn = 0;
        }


        public override void InsertValue()
        {
            UpdateValue(ActiveSheet.LastRowNum);
        }

        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void ReadValue()
        {
            if (ActiveSheet != null)
            {
                //Console.WriteLine(ActiveSheet.SheetName);
                int lastRow = ActiveSheet.LastRowNum;
                for (int i = FirstViewedRow; i <= lastRow; i++)
                {
                    string value1 = "";
                    var row = ActiveSheet.GetRow(i);
                    if (ValueColumn < row.LastCellNum)// -1 это особенность NPOI
                    {
                        value1 = GetCellValue(row.GetCell(ValueColumn));
                    }
                    ProgressValue?.Report((i - FirstViewedRow) / (lastRow - FirstViewedRow) * 100.0);
                    ExchangeValue.Add(value1);
                    double d = (i) / ((double)(ExchangeValue.Count - 1)) * 100.0;
                    ProgressValue?.Report(d);
                }
            }
        }

        public override void UpdateValue()
        {
            UpdateValue(FirstViewedRow);
        }


        /// <summary>
        /// The UpdateValue.
        /// </summary>
        private void UpdateValue(int firstRow)
        {
            int lastRow = Math.Max(ActiveSheet.LastRowNum, ExchangeValue.Count + firstRow);

            for (int i = firstRow; i < lastRow; i++)
            {
                string element = ((List<string>)ExchangeValue).ElementAtOrDefault(i);
                SetCellValue(ActiveSheet, i, ValueColumn, element);
                ProgressValue?.Report(((i - firstRow) / (lastRow - firstRow)) * 100.0);
            }
        }
    }
    */

    public class DictionaryView : ExchangeClass<IDictionary<string, string[]>>
    {
        private readonly MatrixView matrix;
        public DictionaryView(ExchangeOperation exchangeType, string activeSheetName,
            IDictionary<string, string[]> exchangeValue, IProgress<double> progress) :
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

    /*
        /// <summary>
        /// Defines the <see cref="DictionaryView" />.
        /// </summary>
        public class DictionaryViewOLD : ExchangeClass<IDictionary<string, string[]>>
    {
        // обновление по словарю
         // первый столбец ключ
         // второй массив значений соответсвующих этому ключу
         //
        /// <summary>
        /// Gets or sets the KeyColumn.
        /// </summary>
        public int KeyColumn { set; get; } = 0;// столбец в котором находится ключ

        //public Dictionary<string,string[]> ExchangeValue = new Dictionary<string, string[]>();
        /// <summary>
        /// Defines the tmpExchangeValue.
        /// </summary>
        private readonly Dictionary<string, List<string>> tmpExchangeValue = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryView"/> class.
        /// </summary>
        public DictionaryViewOLD(ExchangeOperation exchangeType, string activeSheetName,
            IDictionary<string, string[]> exchangeValue, IProgress<double> progress) :
            base(exchangeType, activeSheetName, progress)
        {
            ExchangeValue = exchangeValue;
        }


        public override void InsertValue()
        {
            var dd = ExchangeValue.SelectMany(x =>x.Value.ToList(),(key,value)=>new string[] { key.Key, value }).ToList();
            Console.WriteLine(dd);
        }


        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void ReadValue()
        {
            if (ActiveSheet != null)
            {
                int lastRow = ActiveSheet.LastRowNum;
                for (int i = FirstViewedRow; i <= lastRow; i++)
                {
                    string keyValue = GetCellValue(ActiveSheet.GetRow(i).GetCell(KeyColumn));
                    string valValue = GetCellValue(ActiveSheet.GetRow(i).GetCell(ValueColumn));
                    if (tmpExchangeValue.ContainsKey(keyValue))
                    {
                        tmpExchangeValue[keyValue].Add(valValue);
                    }
                    else
                    {
                        List<string> tmpList = new()
                        {
                            valValue
                        };
                        tmpExchangeValue[keyValue] = tmpList;
                    }
                    ProgressValue?.Report(((i - FirstViewedRow) / (lastRow - FirstViewedRow)) * 100.0);
                }
                foreach (var list in tmpExchangeValue)
                {
                    ExchangeValue.Add(list.Key, list.Value.ToArray());
                }
            }
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        public override void UpdateValue()
        {
            List<string[]> tmpExchangeValue = new();
            foreach (var keyValue in ExchangeValue)
            {
                foreach (var value in keyValue.Value)
                {
                    string[] tmpValue = new string[2];
                    tmpValue[0] = keyValue.Key;
                    tmpValue[1] = value;
                    tmpExchangeValue.Add(tmpValue);
                }
            }
            int lastRow = Math.Max(ActiveSheet.LastRowNum, tmpExchangeValue.Count + FirstViewedRow);
            for (int i = FirstViewedRow; i <= lastRow; i++)
            {
                var element = tmpExchangeValue.ElementAtOrDefault(i - FirstViewedRow);
                SetCellValue(ActiveSheet, i, KeyColumn, element?.ElementAtOrDefault(0));
                SetCellValue(ActiveSheet, i, ValueColumn, element?.ElementAtOrDefault(1));
                ProgressValue?.Report(((i - FirstViewedRow) / (lastRow - FirstViewedRow)) * 100.0);
            }
        }
    }
    */

    /// <summary>
    /// Defines the <see cref="Wrapper" />.
    /// </summary>
    public class Wrapper : IDisposable
    {
        // To detect redundant calls
        private bool disposed = false;
        internal static ILogger Logger { set; get; }
        // главный класс для обновления
        /// <summary>
        /// Gets or sets the PathToFile.
        /// </summary>
        public readonly string PathToFile;

        private FileStream fileStream; //For disposed. If need to open in other application 

        /// <summary>
        /// Gets or sets the ActiveSheet.
        /// </summary>
        public ISheet ActiveSheet { set; get; } = null;
        public int RowCountActivSheet
        {
            get
            {
                return ActiveSheet.LastRowNum -
                ActiveSheet.FirstRowNum;
            }

        }

        /// <summary>
        /// Gets or sets the ActiveSheetName.
        /// </summary>
        public readonly string ActiveSheetName = "Лист1";

        public string Password { set; get; } = null;

        /// <summary>
        /// Defines the exchangeClass.
        /// </summary>
        public readonly IExchange exchangeClass;

        /// <summary>
        /// Defines the Workbook.
        /// </summary>
        public IWorkbook Workbook;

        /// <summary>
        /// Initializes a new instance of the <see cref="WrapperNpoi"/> class.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        public Wrapper(string pathToFile, IExchange exchangeClass, ILogger logger)
        {
            Logger = logger;
            PathToFile = pathToFile;
            if (exchangeClass != null)
            {
                this.exchangeClass = exchangeClass;
                exchangeClass.Logger = Logger;
                ActiveSheetName = exchangeClass.ActiveSheetName;
            }
            else
            {
                Logger.Error(pathToFile, nameof(exchangeClass));
                throw new ArgumentNullException(nameof(exchangeClass));
            }

        }

        public static string ReturnTechFileName(string predict, string extension)
        {
            int i = 0;
            string rnd = "";
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, predict);
            if (Directory.Exists(dir) == false)
            {
                Directory.CreateDirectory(dir);
            }
            string path;
            do
            {
                path = Path.Combine(dir, $"{predict}{DateTime.Now:yyMMddHHmmss}{rnd}.{extension}");
                i += 1;
                rnd = i.ToString();
            }
            while (File.Exists(path));
            return path;
        }

        public void Exchange()
        {
            switch (exchangeClass.ExchangeOperationEnum)
            {
                case ExchangeOperation.Insert:
                    InsertValue();
                    break;
                case ExchangeOperation.Read:
                    ReadValue();
                    break;
                case ExchangeOperation.Update:
                    UpdateValue();
                    break;
                default:
                    Logger.Error("exchangeClass.ExchangeTypeEnum");
                    throw (new ArgumentOutOfRangeException("exchangeClass.ExchangeTypeEnum"));
            }
        }

        /// <summary>
        /// The ViewWorkbook.
        /// </summary>
        /// <param name="fileMode">The fileMode<see cref="FileMode"/>.</param>
        /// <param name="fileAccess">The fileAccess<see cref="FileAccess"/>.</param>
        /// <param name="exchangeValueFunc">The exchangeValueFunc<see cref="Action"/>.</param>
        /// <param name="addNewWorkbook">The addNewWorkbook<see cref="bool"/>.</param>
        private void ViewWorkbook(FileMode fileMode, FileAccess fileAccess, bool addNewWorkbook, bool closeStream = true, FileShare fileShare=FileShare.ReadWrite)
        {
            if (closeStream == true)
            {
                using FileStream fs = new(PathToFile,
                    fileMode,
                    fileAccess,
                    fileShare);
                Stream tmpStream = fs;
                OpenWorkbookStream(fs, addNewWorkbook);
            }
            else
            {
                fileStream = new(PathToFile,
                fileMode,
                fileAccess,
                fileShare);
                OpenWorkbookStream(fileStream, addNewWorkbook);
            }
        }

        public void OpenWorkbookStream(Stream tmpStream, bool addNewWorkbook)
        {
            FileStream fs = default;
            if (Password == null)
            {}
            else
            {
                NPOI.POIFS.FileSystem.POIFSFileSystem nfs =
                new(fs);
                EncryptionInfo info = new(nfs);
                Decryptor dc = Decryptor.GetInstance(info);
                //bool b = dc.VerifyPassword(Password);
                dc.VerifyPassword(Password);
                tmpStream = dc.GetDataStream(nfs);
            }
            if (addNewWorkbook == true)
            {
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
                if (getValue == false && SheetsCount != 0)
                {
                    ActiveSheet = Workbook.GetSheetAt(0);
                    // поиск в первой если не найдено по наименованию страницы
                }
            }
            exchangeClass.ActiveSheet = ActiveSheet;
            exchangeClass.ExchangeValueFunc();
        }

        /// <summary>
        /// The AddValue.
        /// </summary>
        private void InsertValue()
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
            ViewWorkbook(FileMode.CreateNew, FileAccess.ReadWrite, true, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            Workbook.Write(fs, false);
            fs.Close();
        }

        /// <summary>
        /// The FindValue.
        /// </summary>
        private void ReadValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.ReadValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream,FileShare.Read);
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        private void UpdateValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.UpdateValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            Workbook.Write(fs, false);
            fs.Close();
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        private void OnlyInsertValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.InsertValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            Workbook.Write(fs, false);
            fs.Close();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Освобождаем управляемые ресурсы
                    Logger = null;
                    ActiveSheet = null;
                    Workbook = null;
                    Password = null;
                }
                fileStream?.Close();
            }
            disposed = true;
        }
        // This code added by Visual Basic to
        // correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.
            // Put cleanup code in
            // Dispose(ByVal disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
            GC.Collect();
        }
        ~Wrapper()
        {
            // Do not change this code.
            // Put cleanup code in
            // Dispose(ByVal disposing As Boolean) above.
            Dispose(false);
        }
    }


    /// <summary>
    /// Defines the <see cref="ExchangeExcel" />.
    /// </summary>
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
                Wrapper wrapper = new(pathToFile, listView, null)
                {
                    //ActiveSheetName = sheetName,
                    //exchangeClass = listView
                };
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
                Debug.WriteLine(e.Message);
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
            Wrapper wrapper = new(pathToFile, listView, null)
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
            Wrapper wrapper = new(pathToFile, listView, null)
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
                Wrapper wrapper = new(pathToFile, exchangeClass, null) { };
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
                Wrapper wrapper = new(pathToFile, exchangeClass, null) { };
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
                Wrapper wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else
            {
                throw new TypeUnloadedException("Для указанного типа нет обработчика");
            }
        }
    }

}