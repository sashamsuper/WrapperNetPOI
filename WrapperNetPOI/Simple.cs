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

using Microsoft.Data.Analysis;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace WrapperNetPOI
{
    public static class Simple
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
                Excel.ListView listView = new(ExchangeOperation.Insert, sheetName, values, null)
                {
                    ExchangeValue = values
                };
                Excel.WrapperExcel wrapper = new(pathToFile, listView, null) { };
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
                Task AddValueToExcel = Task.Run(() => AddToExcel(pathToFile, sheetName, values));
            }
            catch (Exception e)
            {
//#if DEBUG
                Wrapper.Logger?.Error(e.Message);
                Wrapper.Logger?.Error(e.StackTrace);
                Debug.WriteLine(e.Message);
//#endif
            }
        }

        public static async Task TaskAddToExcelAsync(string pathToFile, string sheetName, List<string[]> values)
        {
            await Task.Run(() => AddToExcel(pathToFile, sheetName, values));
        }

        /// <summary>
        /// The AddToExcel
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        public static void AddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            Excel.MatrixView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            Excel.WrapperExcel wrapper = new(pathToFile, listView, null)
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
            Excel.ListView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            Excel.WrapperExcel wrapper = new(pathToFile, listView, null)
            {
                //ActiveSheetName = sheetName,
                //exchangeClass = listView
            };
            wrapper.Exchange();
        }

        public static void GetFromExcel<ReturnType>(out ReturnType value, string pathToFile, string sheetName, Excel.Border border = null) where ReturnType : new()
        {
            ReturnType returnValue = new();
            if (returnValue is List<string> rL)
            {
                var exchangeClass = new Excel.ListView(ExchangeOperation.Read, sheetName, rL, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (ReturnType)exchangeClass.ExchangeValue;
                return;
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                var exchangeClass = new Excel.DictionaryView(ExchangeOperation.Read, sheetName, rD, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (ReturnType)exchangeClass.ExchangeValue;
                return;
            }
            else if (returnValue is List<string[]> rM)
            {
                var exchangeClass = new Excel.MatrixView(ExchangeOperation.Read, sheetName, rM, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (ReturnType)exchangeClass.ExchangeValue;
                return;
            }
            else
            {
                var exchangeClass = new Excel.MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (ReturnType)exchangeClass.ExchangeValue;
                return;
            }
            //value = default;
            //return;
        }

        
        public static void GetFromExcelListArray<ReturnType>(out IList<ReturnType[]> value, string pathToFile, string sheetName, Excel.Border border = null) where ReturnType : new()
        {
                var exchangeClass = new Excel.MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = exchangeClass.ExchangeValue;
                return;
        }
        
        
        
        public static void GetFromExcel(out DataFrame value, string pathToFile, string sheetName, Excel.Border border = null, Dictionary<int, Type> header = null, int[] rows = null)
        {
            var exchangeClass = new Excel.DataFrameView(ExchangeOperation.Read, sheetName, null, border);
            if (rows != null)
            {
                exchangeClass.DataHeader = new()
                {
                    Rows = rows
                };
            }
            else
            {
                exchangeClass.DataHeader = new();
            }
            if (header != null)
            {
                exchangeClass.DataHeader.CreateHeaderType(header);
            }
            Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null);
            wrapper.Exchange();
            value = exchangeClass.ExchangeValue;
        }

        /// <summary>
        ///GetFromWord
        /// </summary>
        /// <param name="value"></param>
        /// <param name="pathToFile"></param>
        public static void GetFromWord(out List<Word.TableValue> value, string pathToFile)
        {
            var exchangeClass = new Word.TableView(ExchangeOperation.Read);
            Word.WrapperWord wrapper = new(pathToFile, exchangeClass, null);
            wrapper.Exchange();
            value = exchangeClass.ExchangeValue;
        }

        /// <summary>
        /// The GetFromExcel.
        /// </summary>
        /// <typeparam name="ReturnType">.</typeparam>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <returns>The <see cref="ReturnType"/>.</returns>
        public static ReturnType GetFromExcel<ReturnType>(string pathToFile, string sheetName, Excel.Border border = null) where ReturnType : new()
        {
            GetFromExcel(out ReturnType value, pathToFile, sheetName, border);
            return value;
        }
    }
}