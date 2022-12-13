using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace WrapperNetPOI
{
    
    public enum ExchangeOperation
    {
        Get,
        Add,
        Update
    }

    public interface IExchange
    {
        IProgress<double> ProgressValue { set; get; }
        ILogger Logger { set; get; }
        string ActiveSheetName { set; get; }
        ExchangeOperation ExchangeOperationEnum { set; get; }
        int FirstRow { set; get; }
        ISheet ActiveSheet { set; get; }
        Action ExchangeValueFunc { set; get; }
        bool CloseStream {get;}
        void GetValue();
        void AddValue();
        void UpdateValue();
    }


    /// <summary>
    /// Defines the <see cref="ExchangeClass" />.
    /// </summary>
    public abstract class ExchangeClass<Tout> : IExchange
    {
        

        public virtual bool CloseStream
        {
            get
            {
                return false;
            }
        }
        public IProgress<double> ProgressValue { set; get; }

        public ILogger Logger { set; get; }
        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeClass"/> class.
        /// </summary>
        public ExchangeClass(ExchangeOperation exchange, string activeSheetName)
        {
            ExchangeOperationEnum = exchange;
            ActiveSheetName = activeSheetName;
        }

        public string ActiveSheetName { set; get; }

        public ExchangeOperation ExchangeOperationEnum { set; get; }
        /// <summary>
        /// ActiveSheet
        /// </summary>
        public ISheet ActiveSheet { set; get; }

        // <summary>
        /// $The initial line with which data is entered / viewed$
        /// </summary>
        public int FirstRow { set; get; } = 0;

        /// <summary>
        /// $The initial column from which data is entered/viewed$
        /// </summary>
        public int FirstCol { set; get; } = 0;
        /// <summary>
        /// $Column in which data is entered/viewed$
        /// </summary>
        public int ValueColumn { set; get; } = 1;

        /// <summary>
        /// Gets or sets the ExchangeValue.
        /// </summary>
        public Tout ExchangeValue { set; get; }

        public Action ExchangeValueFunc { set; get; }

        /// <summary>
        /// Maximum column on the right
        /// </summary>
        public int MaxCol { set; get; } = 0;

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
                Logger.Error(e.Message);
                Logger.Error(e.StackTrace);
                return default;
            }

        }



        /// <summary>
        /// The FindValue.
        /// </summary>
        public virtual void GetValue()
        {
            Console.WriteLine("ExchangeGet");
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
        public virtual void AddValue()
        {
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        public virtual void UpdateValue()
        {
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
        public string Path { get; set; }
        //public List<IRow> Rows { set; get; } = new();
        public override bool CloseStream => false;

        public RowsView(ExchangeOperation exchangeType, string activeSheetName, IList<IRow> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;
            ValueColumn = 0;
        }
        public override void GetValue()
        {
            for (int i = FirstRow; i < CountRows; i++)
            {
                ExchangeValue.Add(ActiveSheet.GetRow(i));
            }

        }
        public void AddValue() 
        { 

        }
        
        public override void UpdateValue()
        {
            RowsView rowsView = new(ExchangeOperation.Get, this.ActiveSheetName, new List<IRow>())
            {
                CountRows = this.CountRows
            };
            Wrapper wrapper = new(Path, rowsView);
            wrapper.Exchange();
            /*if (wrapper.ActiveSheet is XSSFSheet xsSheet)
            {
                /*CellCopyPolicy cellCopyPolicy = new ();
                if (rowsView.ExchangeValue is List<XSSFRow> xSSFRows)
                {
                    xsSheet.CopyRows(xSSFRows, FirstRow, cellCopyPolicy);
                }
            }
            else 
            {*/
                for (int i=0;i<CountRows;i++)
                {
                    SheetUtil.CopyRow(rowsView.ActiveSheet, i, wrapper.ActiveSheet, i);
                }
            
        }
    }




    /// <summary>
    /// Defines the <see cref="MatrixView" />.
    /// </summary>
    public class MatrixView : ExchangeClass<IList<string[]>>
    {
        /// <summary>
        /// Defines the ExchangeValue1.
        /// </summary>
        //public List<string[]> ExchangeValue1 = new List<string[]>();

        // занесение спика массива=строкам
        /// <summary>
        /// Initializes a new instance of the <see cref="MatrixView"/> class.
        /// </summary>
        public MatrixView(ExchangeOperation exchangeType, string activeSheetName, IList<string[]> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;

            ValueColumn = 0;
        }

                
        /// <summary>
        /// The AddValue.
        /// </summary>
        /// <param name="startRow">The startRow<see cref="int"/>.</param>
        public void AddValue(int startRow)
        {
            for (int i = startRow; i < ExchangeValue.Count; i++)
            {
                IRow row = ActiveSheet.CreateRow(i);
                for (int j = 0; j < ExchangeValue[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(ExchangeValue[i][j]);
                }
                double d = (i) / ((double)(ExchangeValue.Count - 1)) * 100.0;
                ProgressValue?.Report(d);
            }
        }

        /// <summary>
        /// The AddValue.
        /// </summary>
        public override void AddValue()
        {
            for (int i = 0; i < ExchangeValue.Count; i++)
            {
                //IRow row = ActiveSheet.CreateRow(i);
                for (int j = 0; j < ExchangeValue[i].Length; j++)
                {
                    SetCellValue(ActiveSheet, i, j, ExchangeValue[i][j]);
                }
                double d = (i) / ((double)(ExchangeValue.Count - 1)) * 100.0;
                ProgressValue?.Report(d);
            }
        }



        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void GetValue()
        {
            if (ActiveSheet != null)
            {
                int lastRow = ActiveSheet.LastRowNum;
                for (int i = FirstRow; i <= lastRow; i++)
                {
                    var row = ActiveSheet.GetRow(i);
                    List<string> tmpListString = new();
                    if (row != null)
                    {
                        var lastCol = row.LastCellNum;
                        for (int ValueColumn = FirstCol; ValueColumn <= lastCol; ValueColumn++)
                        {
                            ICell cell = row.GetCell(ValueColumn);
                            if (ValueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                            {
                                tmpListString.Add(GetCellValue(cell));
                            }
                            ProgressValue?.Report((i - FirstRow) / (lastRow - FirstRow) * 100.0);
                        }
                    }
                    ExchangeValue.Add(tmpListString.ToArray());
                }
            }
        }
    }

    /// <summary>
    /// Defines the <see cref="ListView" />.
    /// </summary>
    public class ListView : ExchangeClass<IList<string>>
    {
        // обновление по листу
        /// <summary>
        /// Initializes a new instance of the <see cref="ListView"/> class.
        /// </summary>
        public ListView(ExchangeOperation exchangeType, string activeSheetName, IList<string> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;
            ValueColumn = 0;
        }

        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void GetValue()
        {
            if (ActiveSheet != null)
            {
                Console.WriteLine(ActiveSheet.SheetName);
                int lastRow = ActiveSheet.LastRowNum;
                for (int i = FirstRow; i <= lastRow; i++)
                {
                    string value1 = "";
                    var row = ActiveSheet.GetRow(i);
                    if (ValueColumn <= row.LastCellNum - 1)// -1 это особенность NPOI
                    {
                        value1 = GetCellValue(row.GetCell(ValueColumn));
                    }
                    ProgressValue?.Report((i - FirstRow) / (lastRow - FirstRow) * 100.0);
                    ExchangeValue.Add(value1);
                }
            }
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        public override void UpdateValue()
        {
            int lastRow = Math.Max(ActiveSheet.LastRowNum, ExchangeValue.Count + FirstRow);

            for (int i = FirstRow; i <= lastRow; i++)
            {
                string element = ((List<string>)ExchangeValue).ElementAtOrDefault(i);
                SetCellValue(ActiveSheet, i, ValueColumn, element);
                ProgressValue?.Report(((i - FirstRow) / (lastRow - FirstRow)) * 100.0);
            }
        }
    }

    /// <summary>
    /// Defines the <see cref="DictionaryView" />.
    /// </summary>
    public class DictionaryView : ExchangeClass<IDictionary<string, string[]>>
    {
        /* обновление по словарю
         * первый столбец ключ
         * второй массив значений соответсвующих этому ключу
         */
        /// <summary>
        /// Gets or sets the KeyColumn.
        /// </summary>
        public int KeyColumn { set; get; } = 0;// столбец в котором находится ключ

        //public Dictionary<string,string[]> ExchangeValue = new Dictionary<string, string[]>();
        /// <summary>
        /// Defines the tmpExchangeValue.
        /// </summary>
        private Dictionary<string, List<string>> tmpExchangeValue = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryView"/> class.
        /// </summary>
        public DictionaryView(ExchangeOperation exchangeType, string activeSheetName, IDictionary<string, string[]> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;
        }

        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void GetValue()
        {
            if (ActiveSheet != null)
            {
                int lastRow = ActiveSheet.LastRowNum;
                for (int i = FirstRow; i <= lastRow; i++)
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
            int lastRow = Math.Max(ActiveSheet.LastRowNum, tmpExchangeValue.Count + FirstRow);
            for (int i = FirstRow; i <= lastRow; i++)
            {
                var element = tmpExchangeValue.ElementAtOrDefault(i - FirstRow);
                SetCellValue(ActiveSheet, i, KeyColumn, element?.ElementAtOrDefault(0));
                SetCellValue(ActiveSheet, i, ValueColumn, element?.ElementAtOrDefault(1));
                ProgressValue?.Report(((i - FirstRow) / (lastRow - FirstRow)) * 100.0);
            }
        }
    }

    /// <summary>
    /// Defines the <see cref="Wrapper" />.
    /// </summary>
    public class Wrapper:IDisposable
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
        public Wrapper(string pathToFile, IExchange exchangeClass)
        {

            var pathLog = Wrapper.ReturnTechFileName("Log", "log");
            Logger = new LoggerConfiguration()
             .MinimumLevel.Verbose() // ставим минимальный уровень в Verbose для теста, по умолчанию стоит Information 
                                     //.WriteTo.Console()  // выводим данные на консоль
             .WriteTo.File(pathLog) // а также пишем лог файл, разбивая его по дате
             .CreateLogger();
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
                case ExchangeOperation.Add:
                    AddValue();
                    break;
                case ExchangeOperation.Get:
                    GetValue();
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
        private void ViewWorkbook(FileMode fileMode, FileAccess fileAccess, bool addNewWorkbook, bool closeStream = true)
        {
            if (closeStream == true)
            {
                using (FileStream fs = new(PathToFile,
                    fileMode,
                    fileAccess,
                    FileShare.ReadWrite))
                {
                    Stream tmpStream = fs;
                    OpenWorkbookStream(fs, addNewWorkbook);
                }
            }
            else
            {
                FileStream fs = new(PathToFile,
                fileMode,
                fileAccess,
                FileShare.ReadWrite);
                OpenWorkbookStream(fs, addNewWorkbook);
            }


        }

        public void OpenWorkbookStream(Stream tmpStream, bool addNewWorkbook)
        {
            FileStream fs = default;
            if (Password == null)
            {

            }
            else
            {
                NPOI.POIFS.FileSystem.POIFSFileSystem nfs =
                new(fs);
                EncryptionInfo info = new(nfs);
                Decryptor dc = Decryptor.GetInstance(info);
                bool b = dc.VerifyPassword(Password);
                tmpStream = dc.GetDataStream(nfs);
            }

            if (addNewWorkbook == true)
            {
                Workbook = new XSSFWorkbook();
            }
            else
            {
                Workbook = WorkbookFactory.Create(tmpStream);
            }

            int SheetsCount = Workbook.NumberOfSheets;
            bool getValue = false;
            for (int i = 0; i < Workbook.NumberOfSheets; i++)
            {
                if (Workbook.GetSheetAt(i).SheetName == ActiveSheetName)
                {
                    if (Workbook.GetSheet(ActiveSheetName) is XSSFSheet activeSheet)
                    {
                        ActiveSheet = activeSheet;
                        getValue = true;
                    }
                    else if (Workbook.GetSheet(ActiveSheetName) is HSSFSheet activeSheet2)
                    {
                        ActiveSheet = activeSheet2;
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
            if (ActiveSheet == null)
                if (addNewWorkbook)
                {
                    Workbook.CreateSheet(ActiveSheetName);
                    ActiveSheet = (XSSFSheet)Workbook.GetSheet(ActiveSheetName);
                }
            exchangeClass.ActiveSheet = ActiveSheet;
            //Console.WriteLine(exchangeClass.ExchangeValueFunc);
            exchangeClass.ExchangeValueFunc();
        }

        /// <summary>
        /// The AddValue.
        /// </summary>
        private void AddValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.AddValue;
            ViewWorkbook(FileMode.CreateNew, FileAccess.ReadWrite, true,exchangeClass.CloseStream);
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
        private void GetValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.GetValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false,exchangeClass.CloseStream);
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        private void UpdateValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.UpdateValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, true,exchangeClass.CloseStream);
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
                    Workbook=null;
                    Password=null;
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
                ListView listView = new(ExchangeOperation.Add, sheetName, values)
                {
                    ExchangeValue = values
                };
                Wrapper wrapper = new(pathToFile, listView)
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

            //AddValueToExcel.Start();
        }

        public static async Task TaskAddToExcelAsync(string pathToFile, string sheetName, List<string[]> values)
        {
            await Task.Run(() =>
            {
                AddToExcel(pathToFile, sheetName, values);
            });
            //AddValueToExcel.Start();
        }







        /// <summary>
        /// The AddToExcel
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        public static void AddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            MatrixView listView = new(ExchangeOperation.Add, sheetName, values)
            {
                ExchangeValue = values
            };
            Wrapper wrapper = new(pathToFile, listView)
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
            ListView listView = new(ExchangeOperation.Add, sheetName, values)
            {
                ExchangeValue = values
            };
            Wrapper wrapper = new(pathToFile, listView)
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

            //IExchange exchangeClass;
            ReturnType returnValue = new();
            if (returnValue is List<string> rL)
            {
                var exchangeClass = new ListView(ExchangeOperation.Get, sheetName, rL)
                {
                    FirstCol = firstCol,
                    FirstRow = firstRow
                };
                Wrapper wrapper = new(pathToFile, exchangeClass) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                var exchangeClass = new DictionaryView(ExchangeOperation.Get, sheetName, rD)
                {
                    FirstCol = firstCol,
                    FirstRow = firstRow
                };
                Wrapper wrapper = new(pathToFile, exchangeClass) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else if (returnValue is List<string[]> rM)
            {
                var exchangeClass = new MatrixView(ExchangeOperation.Get, sheetName, rM)
                {
                    FirstCol = firstCol,
                    FirstRow = firstRow
                };
                Wrapper wrapper = new(pathToFile, exchangeClass) { };
                wrapper.Exchange();
                return (ReturnType)exchangeClass.ExchangeValue;
            }
            else
            {
                throw new TypeUnloadedException("Для указанного типа нет обработчика");
            }


            //Console.WriteLine(exchangeClass.ExchangeValue.Count);
            //Console.WriteLine(exchangeClass.ExchangeValue.Count);
            //string json = JsonSerializer.Serialize(wrapper);
            //Console.WriteLine(exchangeClass);

            //RoslynSoftDebugger.HitBreakpoint(); 
            //RoslynSoftDebugger.Debug(()=> n);
            //RoslynSoftDebugger.Debug();
        }
    }

}