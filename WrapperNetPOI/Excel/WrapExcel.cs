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
using System.IO;
namespace WrapperNetPOI.Excel
{
    public class WrapperExcel : Wrapper
    {
        public WrapperExcel(string pathToFile, IExchangeExcel ExcelExchange, ILogger logger = null) :
        base(pathToFile, ExcelExchange, logger)
        { }
        protected override void InsertValue()
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
            ExcelExchange.ExchangeValueFunc = ExcelExchange.InsertValue;
            ViewFile(FileMode.CreateNew, FileAccess.ReadWrite, true, ExcelExchange.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)ExcelExchange).Workbook.Write(fs, false);
            fs.Close();
        }
        protected override void ReadValue()
        {
            ExcelExchange.ExchangeValueFunc = ExcelExchange.ReadValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, ExcelExchange.CloseStream, FileShare.Read);
        }
        protected override void UpdateValue()
        {
            ExcelExchange.ExchangeValueFunc = ExcelExchange.UpdateValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, ExcelExchange.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)ExcelExchange).Workbook.Write(fs, false);
            fs.Close();
        }
        protected override void DeleteValue()
        {
            ExcelExchange.ExchangeValueFunc = ExcelExchange.DeleteValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, ExcelExchange.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)ExcelExchange).Workbook.Write(fs, false);
            fs.Close();
        }
        private void OnlyInsertValue()
        {
            ExcelExchange.ExchangeValueFunc = ExcelExchange.InsertValue;
            ViewFile(FileMode.Open, FileAccess.Read, false, ExcelExchange.CloseStream);
            using FileStream fs = new(PathToFile,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.ReadWrite);
            ((IExchangeExcel)ExcelExchange).Workbook.Write(fs, false);
            fs.Close();
        }
    }
}