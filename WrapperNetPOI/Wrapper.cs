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
    /// <summary>
    /// The exchanges operations.
    /// </summary>
    public enum ExchangeOperation
    {
        Insert,
        Read,
        Update,
        Delete
    }

    /// <summary>
    /// The exchange interface.
    /// </summary>
    public interface IExchange
    {
        //IWorkbook Workbook {set;get;}
        /// <summary>
        /// Gets or sets the progress value.
        /// </summary>
        IProgress<int> ProgressValue { set; get; }

        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        ILogger Logger { set; get; }
        /// <summary>
        /// Gets or sets the exchange operation enum.
        /// </summary>
        ExchangeOperation ExchangeOperationEnum { set; get; }
        /// <summary>
        /// Gets or sets the exchange value func.
        /// </summary>
        Action ExchangeValueFunc { set; get; }
        /// <summary>
        /// Gets or sets a value indicating whether close stream.
        /// </summary>
        bool CloseStream { get; set; }

        /// <summary>
        /// Get internally object.
        /// </summary>
        /// <param name="fs">The fs.</param>
        /// <param name="addNew">If true, add new.</param>
        void GetInternallyObject(Stream fs, bool addNew);

        /// <summary>
        /// Reads the value.
        /// </summary>
        void ReadValue();

        /// <summary>
        /// Inserts the value.
        /// </summary>
        void InsertValue();

        /// <summary>
        /// Update the value.
        /// </summary>
        void UpdateValue();

        /// <summary>
        /// Deletes the value.
        /// </summary>
        void DeleteValue();
    }

    /// <summary>
    /// The wrapper.
    /// </summary>
    public abstract class Wrapper : IDisposable //Main class
    {
        // To detect redundant calls
        /// <summary>
        /// The disposed.
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        internal static ILogger Logger { set; get; }

        ///<summary>
        /// /// Gets or sets the PathToFile.
        /// </summary>
        public readonly string PathToFile;

        /// <summary>
        /// File stream.
        /// </summary>
        protected FileStream fileStream; //For disposed. If need to open in other application

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
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

        /// <summary>
        /// Return tech file name.
        /// </summary>
        /// <param name="predict">The predict.</param>
        /// <param name="extension">The extension.</param>
        /// <returns>A string</returns>
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
            } while (File.Exists(path));
            return path;
        }

        /// <summary>
        /// View the file.
        /// </summary>
        /// <param name="fileMode">The file mode.</param>
        /// <param name="fileAccess">The file access.</param>
        /// <param name="addNew">If true, add new.</param>
        /// <param name="closeStream">If true, close stream.</param>
        /// <param name="fileShare">The file share.</param>
        protected void ViewFile(
            FileMode fileMode,
            FileAccess fileAccess,
            bool addNew,
            bool closeStream = true,
            FileShare fileShare = FileShare.ReadWrite
        )
        {
            if (closeStream)
            {
                using FileStream fs = new(PathToFile, fileMode, fileAccess, fileShare);
                Stream tmpStream = fs;
                exchangeClass.GetInternallyObject(fs, addNew);
            }
            else // Apparently it's useless for NPOI
            {
                fileStream = new(PathToFile, fileMode, fileAccess, fileShare);
                exchangeClass.GetInternallyObject(fileStream, addNew);
            }
        }

        /// <summary>
        /// Inserts the value.
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        protected virtual void InsertValue()
        {
            throw new NotImplementedException("InsertValue");
        }

        /// <summary>
        /// Reads the value.
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        protected virtual void ReadValue()
        {
            throw new NotImplementedException("ReadValue");
        }

        /// <summary>
        /// Update the value.
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        protected virtual void UpdateValue()
        {
            throw new NotImplementedException("UpdateValue");
        }

        /// <summary>
        /// Deletes the value.
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        protected virtual void DeleteValue()
        {
            throw new NotImplementedException("DeleteValue");
        }

        /// <summary>
        /// 
        /// </summary>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing">If true, disposing.</param>
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
        /// <summary>
        /// 
        /// </summary>
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
