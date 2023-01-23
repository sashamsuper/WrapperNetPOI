
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
using Serilog;
using System;
using System.IO;

namespace WrapperNetPOI
{

    public class WrapperExcel : Wrapper
    {
        /// <summary>
        /// Gets or sets the ActiveSheet.
        /// </summary>
        //public ISheet ActiveSheet { set; get; } = null;

        /// <summary>
        /// Gets or sets the ActiveSheetName.
        /// </summary>

        /* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
        До:
                //public readonly string ActiveSheetName = "List1";

                public WrapperExcel(string pathToFile, IExchangeExcel exchangeClass, ILogger logger = null):
        После:
                //public readonly string ActiveSheetName = "List1";

                public WrapperExcel(string pathToFile, IExchangeExcel exchangeClass, ILogger logger = null):
        */
        //public readonly string ActiveSheetName = "List1";

        public WrapperExcel(string pathToFile, IExchangeExcel exchangeClass, ILogger logger = null) :
        base(pathToFile, exchangeClass, logger)
        {
            //ActiveSheetName = exchangeClass.ActiveSheetName;
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
            ViewFile(FileMode.CreateNew, FileAccess.ReadWrite, true, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)exchangeClass).Workbook.Write(fs, false);
            fs.Close();
        }

        private void ReadValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.ReadValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream, FileShare.Read);
        }

        private void UpdateValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.UpdateValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)exchangeClass).Workbook.Write(fs, false);
            fs.Close();
        }

        private void OnlyInsertValue()
        {
            exchangeClass.ExchangeValueFunc = exchangeClass.InsertValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, exchangeClass.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)exchangeClass).Workbook.Write(fs, false);
            fs.Close();
        }

        /* Необъединенное слияние из проекта "WrapperNetPOI (net6.0)"
        До:
            }




            public abstract class Wrapper : IDisposable //Main class
        После:
            }




            public abstract class Wrapper : IDisposable //Main class
        */
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
        //public IWorkbook Workbook;

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

        protected void ViewFile(FileMode fileMode, FileAccess fileAccess, bool addNew, bool closeStream = true, FileShare fileShare = FileShare.ReadWrite)
        {
            if (closeStream == true)
            {
                using FileStream fs = new(PathToFile,
                    fileMode,
                    fileAccess,
                    fileShare);
                Stream tmpStream = fs;
                exchangeClass.GetInternallyObject(fs, addNew);
            }
            else
            {
                fileStream = new(PathToFile,
                fileMode,
                fileAccess,
                fileShare);
                exchangeClass.GetInternallyObject(fileStream, addNew);
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
                    //Workbook = null;
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