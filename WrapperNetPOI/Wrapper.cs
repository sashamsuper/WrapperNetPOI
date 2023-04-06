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

using Serilog;
using System;
using System.IO;

namespace WrapperNetPOI
{
    public enum ExchangeOperation
    {
        Insert,
        Read,
        Update,
        Delete
    }
    public interface IExchange
    {
        //IWorkbook Workbook {set;get;}
        IProgress<int> ProgressValue { set; get; }

        ILogger Logger { set; get; }
        ExchangeOperation ExchangeOperationEnum { set; get; }
        Action ExchangeValueFunc { set; get; }
        bool CloseStream { get; set; }

        void GetInternallyObject(Stream fs, bool addNew);

        void ReadValue();

        void InsertValue();

        void UpdateValue();

        void DeleteValue();
    }
    public abstract class Wrapper : IDisposable //Main class
    {
        // To detect redundant calls
        private bool disposed = false;

        internal static ILogger Logger { set; get; }

        ///<summary>
        /// /// Gets or sets the PathToFile.
        /// </summary>
        public readonly string PathToFile;

        protected FileStream fileStream; //For disposed. If need to open in other application

        public string Password { set; get; } = null;

        /// <summary>
        /// Defines the exchangeClass.
        /// </summary>
        public readonly IExchange exchangeClass;

        /// <summary>
        /// Initializes a new instance of the <see cref="WrapperNpoi"/> class.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        protected Wrapper(string pathToFile, IExchange exchangeClass, ILogger logger = null)
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
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, predict);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            string path;
            do
            {
                path = Path.Combine(dir, $"{predict}{DateTime.Now:yyMMddHHmmss}{i}.{extension}");
                i++;
            }
            while (File.Exists(path));
            return path;
        }

        protected void ViewFile(FileMode fileMode, FileAccess fileAccess, bool addNew, bool closeStream = true, FileShare fileShare = FileShare.ReadWrite)
        {
            if (closeStream)
            {
                using FileStream fs = new(PathToFile,
                    fileMode,
                    fileAccess,
                    fileShare);
                Stream tmpStream = fs;
                exchangeClass.GetInternallyObject(fs, addNew);
            }
            else // Apparently it's useless for NPOI
            {
                fileStream = new(PathToFile,
                fileMode,
                fileAccess,
                fileShare);
                exchangeClass.GetInternallyObject(fileStream, addNew);
            }
        }

        protected virtual void InsertValue()
        {
            throw new NotImplementedException("InsertValue");
        }

        protected virtual void ReadValue()
        {
            throw new NotImplementedException("ReadValue");
        }

        protected virtual void UpdateValue()
        {
            throw new NotImplementedException("UpdateValue");
        }

        protected virtual void DeleteValue()
        {
            throw new NotImplementedException("DeleteValue");
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

                case ExchangeOperation.Delete:
                    DeleteValue();
                    break;

                default:
                    Logger.Error("exchangeClass.ExchangeTypeEnum");
                    throw (new ArgumentOutOfRangeException("exchangeClass.ExchangeTypeEnum"));
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