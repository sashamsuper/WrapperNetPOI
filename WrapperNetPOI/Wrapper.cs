
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
    
    public class WrapperExcel:Wrapper
    {
        /// <summary>
        /// Gets or sets the ActiveSheet.
        /// </summary>
        public ISheet ActiveSheet { set; get; } = null;


        /// <summary>
        /// Gets or sets the ActiveSheetName.
        /// </summary>
        public readonly string ActiveSheetName = "List1";
        
        public WrapperExcel(string pathToFile, IExchange exchangeClass, ILogger logger = null):
        base(pathToFile, exchangeClass, logger)
        {
            ActiveSheetName = exchangeClass.ActiveSheetName;
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

        private void ViewWorkbook(FileMode fileMode, FileAccess fileAccess, bool addNewWorkbook, bool closeStream = true, FileShare fileShare = FileShare.ReadWrite)
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
            { }
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
                    // search first if not found
                }
            }
            exchangeClass.ActiveSheet = ActiveSheet;
            exchangeClass.ExchangeValueFunc();
        }

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

        private void ReadValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.ReadValue;
            ViewWorkbook(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream, FileShare.Read);
        }

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
    }
    
    
    
    public abstract class Wrapper : IDisposable //Main class
    {
        // To detect redundant calls
        private bool disposed = false;
        internal static ILogger Logger { set; get; }

        /// Gets or sets the PathToFile.
        /// </summary>
        public readonly string PathToFile;

        protected FileStream fileStream; //For disposed. If need to open in other application 

        
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
        public Wrapper(string pathToFile, IExchange exchangeClass, ILogger logger = null)
        {
            Logger = logger;
            PathToFile = pathToFile;
            if (exchangeClass != null)
            {
                this.exchangeClass = exchangeClass;
                exchangeClass.Logger = Logger;
                
            }
            else
            {
                Logger.Error(pathToFile, nameof(exchangeClass));
                throw new ArgumentNullException(nameof(exchangeClass));
            }

        }

        

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Освобождаем управляемые ресурсы
                    
                    Logger = null;
                    
                    
                    
                    //ActiveSheet = null;
                    
                    
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
}