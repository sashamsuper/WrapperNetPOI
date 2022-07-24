namespace WrappperNPOI
{
    using NPOI.HSSF.UserModel;
    using NPOI.POIFS.Crypt;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    public enum ExchangeType
    {
        Get,
        Add,
        Update
    }


    /// <summary>
    /// Defines the <see cref="ExchangeClass" />.
    /// </summary>
    public abstract class ExchangeClass
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeClass"/> class.
        /// </summary>
        public ExchangeClass(ExchangeType exchange, string activeSheetName)
        {
            ExchangeTypeEnum = exchange;
            ActiveSheetName = activeSheetName;
        }

        public readonly string ActiveSheetName;

        public readonly ExchangeType ExchangeTypeEnum;
        /// <summary>
        /// Gets or sets the ActiveSheet.
        /// </summary>
        public ISheet ActiveSheet { set; get; }

        // <summary>
        /// $Начальная строка с которой идет ввод/просмотр данных$
        /// </summary>
        public int FirstRow { set; get; } = 0;// начальная строка с которой идет ввод/просмотр данных

        /// <summary>
        /// $Начальный столбец с которого идет ввод/просмотр данных$
        /// </summary>
        public int FirstCol { set; get; } = 0;// начальный столбец с которого идет ввод/просмотр данных

        /// <summary>
        /// $Столбец в которого вводятся/просматриваются данные$
        /// </summary>
        public int ValueColumn { set; get; } = 1;// столбец в который вводятся/просматриваются данные

        /// <summary>
        /// Gets or sets the ExchangeValue.
        /// </summary>
        public dynamic ExchangeValue { set; get; }

        public Action ExchangeValueFunc { set; get; }

        /// <summary>
        /// $Максимальное столбец справа$
        /// </summary>
        public int MaxCol { set; get; } = 0;// максимальное столбец справа;

        /// <summary>
        /// $Возврат даты в формате dd.mm.yyyy$
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

        public static string GetCellValue(ICell cell)
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
                    returnValue = ExchangeClass.
                        ReturnStringDate(cell.DateCellValue);
                }
                else if (cell?.CellType == CellType.Numeric
                && cell.ToString()?.Split('.').Length >= 3)
                {
                    returnValue = ExchangeClass.ReturnStringDate(cell.DateCellValue);
                }
                else if (cell?.CellType == CellType.Formula)
                {
                    if (cell?.CachedFormulaResultType == CellType.Numeric
                  && cell.NumericCellValue > 36526 &&
                  cell.NumericCellValue < 47484)
                    {
                        returnValue = ExchangeClass.
                            ReturnStringDate(cell.DateCellValue);
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

    /// <summary>
    /// Defines the <see cref="MatrixView" />.
    /// </summary>
    public class MatrixView : ExchangeClass
    {
        /// <summary>
        /// Defines the ExchangeValue1.
        /// </summary>
        //public List<string[]> ExchangeValue1 = new List<string[]>();

        // занесение спика массива=строкам
        /// <summary>
        /// Initializes a new instance of the <see cref="MatrixView"/> class.
        /// </summary>
        public MatrixView(ExchangeType exchangeType, string activeSheetName, IList<string[]> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;

            ValueColumn = 0;
        }

        // ExchangeValue - значение куда записываются выходные данные
        // и куда записваются данные которые надо записать в файл 
        //public new int FirstRow { set; get; } = 0; // начальная строка с которой идет ввод/просмотр данных
        //public new int ActiveColumn { set; get; } = 0; // столбец в который вводятся/просматриваются данные
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
    public class ListView : ExchangeClass
    {
        // обновление по листу
        /// <summary>
        /// Initializes a new instance of the <see cref="ListView"/> class.
        /// </summary>
        public ListView(ExchangeType exchangeType, string activeSheetName, IList<string> exchangeValue) : base(exchangeType, activeSheetName)
        {
            ExchangeValue = exchangeValue;
            ValueColumn = 0;
        }

        // ExchangeValue - значение куда записываются выходные данные
        // и куда записваются данные которые надо записать в файл 
        //public new int FirstRow { set; get; } = 0; // начальная строка с которой идет ввод/просмотр данных
        //public new int ActiveColumn { set; get; } = 0; // столбец в который вводятся/просматриваются данные
        /// <summary>
        /// The GetValue.
        /// </summary>
        public override void GetValue()
        {
            Console.WriteLine("blinn null");
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
            }
        }
    }

    /// <summary>
    /// Defines the <see cref="DictionaryView" />.
    /// </summary>
    public class DictionaryView : ExchangeClass
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
        public DictionaryView(ExchangeType exchangeType, string activeSheetName, IDictionary<string, string[]> exchangeValue) : base(exchangeType, activeSheetName)
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
                    //string keyValue = ActiveSheet.GetRow(i).GetCell(KeyColumn).ToString();
                    //string valValue = ActiveSheet.GetRow(i).GetCell(ValueColumn).ToString();
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
            }
        }
    }

    /// <summary>
    /// Defines the <see cref="WrapperNpoi" />.
    /// </summary>
    public class WrapperNpoi
    {
        // главный класс для обновления
        /// <summary>
        /// Gets or sets the PathToFile.
        /// </summary>
        public readonly string PathToFile;

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
        public readonly ExchangeClass exchangeClass;

        /// <summary>
        /// Defines the Workbook.
        /// </summary>
        public IWorkbook Workbook;

        /// <summary>
        /// Initializes a new instance of the <see cref="WrapperNpoi"/> class.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        public WrapperNpoi(string pathToFile, ExchangeClass exchangeClass)
        {
            PathToFile = pathToFile;
            //ActiveSheetName=exchangeClass.activeSheetName;
            if (exchangeClass != null)
            {
                this.exchangeClass = exchangeClass;

                ActiveSheetName = exchangeClass.ActiveSheetName;
            }
            else
            {
                throw new ArgumentNullException(nameof(exchangeClass));
            }

        }

        public void Exchange()
        {
            switch (exchangeClass.ExchangeTypeEnum)
            {
                case ExchangeType.Add:
                    Console.WriteLine("Add");
                    AddValue();
                    break;
                case ExchangeType.Get:
                    Console.WriteLine("Find");
                    GetValue();
                    break;
                case ExchangeType.Update:
                    UpdateValue();
                    break;
                default:
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
        private void ViewWorkbook(FileMode fileMode, FileAccess fileAccess, bool addNewWorkbook)
        {
            using FileStream fs = new(PathToFile,
                    fileMode,
                    fileAccess,
                    FileShare.ReadWrite);
            Stream tmpStream = fs;
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

            Workbook = WorkbookFactory.Create(tmpStream);
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
            Console.WriteLine(exchangeClass.ExchangeValueFunc);
            exchangeClass.ExchangeValueFunc();
        }

        /// <summary>
        /// The AddValue.
        /// </summary>
        private void AddValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.AddValue;
            ViewWorkbook(FileMode.CreateNew, FileAccess.ReadWrite, true);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            Workbook.Write(fs);
            fs.Close();
        }

        /// <summary>
        /// The FindValue.
        /// </summary>
        private void GetValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.GetValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false);
        }

        /// <summary>
        /// The UpdateValue.
        /// </summary>
        private void UpdateValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.UpdateValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, true);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            Workbook.Write(fs);
            fs.Close();
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
                ListView listView = new(ExchangeType.Add, sheetName, values)
                {
                    ExchangeValue = values
                };
                WrapperNpoi wrapper = new(pathToFile, listView)
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
            Task AddValueToExcel = Task.Run(() =>
            {
                TaskAddToExcel(pathToFile, sheetName, values);
            });
            AddValueToExcel.Start();
        }

        /// <summary>
        /// The AddToExcel
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        public static void AddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            MatrixView listView = new(ExchangeType.Add, sheetName, values)
            {
                ExchangeValue = values
            };
            WrapperNpoi wrapper = new(pathToFile, listView)
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
            ListView listView = new(ExchangeType.Add, sheetName, values)
            {
                ExchangeValue = values
            };
            WrapperNpoi wrapper = new(pathToFile, listView)
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
        public static ReturnType GetFromExcel<ReturnType>(string pathToFile, string sheetName,int firstRow=0, int firstCol=0) where ReturnType : new()
        {
            ExchangeClass exchangeClass;
            ReturnType returnValue = new();
            if (returnValue is List<string> rL)
            {
                exchangeClass = new ListView(ExchangeType.Get, sheetName, rL)
                {
                    FirstCol=firstCol,
                    FirstRow=firstRow
                    //ExchangeValue = returnValue
                };
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                exchangeClass = new DictionaryView(ExchangeType.Get, sheetName, rD)
                {
                    FirstCol=firstCol,
                    FirstRow=firstRow
                    //ExchangeValue = returnValue
                };
            }
            else if (returnValue is List<string[]> rM)
                exchangeClass = new MatrixView(ExchangeType.Get, sheetName, rM)
                {
                   FirstCol=firstCol,
                   FirstRow=firstRow
                   //ExchangeValue = returnValue
                };
            else
            {
                throw new TypeUnloadedException("Для указанного типа нет обработчика");
            }

            WrapperNpoi wrapper = new(pathToFile, exchangeClass)
            {
                //ActiveSheetName = sheetName,
                //exchangeClass = exchangeClass
            };
            Console.WriteLine(exchangeClass.ExchangeValue.Count);
            wrapper.Exchange();
            Console.WriteLine(exchangeClass.ExchangeValue.Count);
            //string json = JsonSerializer.Serialize(wrapper);
            Console.WriteLine(exchangeClass);

            //RoslynSoftDebugger.HitBreakpoint(); 
            //RoslynSoftDebugger.Debug(()=> n);
            //RoslynSoftDebugger.Debug();
            return wrapper.exchangeClass.ExchangeValue;
        }
    }

}