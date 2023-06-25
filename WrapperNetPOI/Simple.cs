using System.Security.Cryptography;
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
using WrapperNetPOI.Excel;
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
        private static void TaskAddToExcel(string pathToFile, string sheetName, List<string> values)
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
        private static void TaskAddToExcel(string pathToFile, string sheetName, List<string[]> values)
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
        private static async Task TaskAddToExcelAsync(string pathToFile, string sheetName, List<string[]> values)
        {
            await Task.Run(() => AddToExcel(pathToFile, sheetName, values));
        }
        /// <summary>
        /// The AddToExcel
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string[]}"/>.</param>
        private static void AddToExcel(string pathToFile, string sheetName, List<string[]> values)
        {
            
            Excel.MatrixView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            Excel.WrapperExcel wrapper = new(pathToFile, listView, null)
            {
                //ActiveSheetName = sheetName,
                //ExcelExchange = listView
            };
            wrapper.Exchange();
        }
        /// <summary>
        /// The AddToExcel.
        /// </summary>
        /// <param name="pathToFile">The pathToFile<see cref="string"/>.</param>
        /// <param name="sheetName">The sheetName<see cref="string"/>.</param>
        /// <param name="values">The values<see cref="List{string}"/>.</param>
        private static void AddToExcel(string pathToFile, string sheetName, List<string> values)
        {
            Excel.ListView listView = new(ExchangeOperation.Insert, sheetName, values, null)
            {
                ExchangeValue = values
            };
            Excel.WrapperExcel wrapper = new(pathToFile, listView, null)
            {
            };
            wrapper.Exchange();
        }

        public static void InsertToExcel<TInsert>(TInsert value, string pathToFile, string sheetName="Sheet1", Border border=null)
        {
            Action action = value switch
            {
                List<String> when value is List<string> listStr => new Action(() =>
                {
                    ListView listView = new(ExchangeOperation.Insert, sheetName, listStr, null)
                    {
                        ExchangeValue= listStr
                    };
                    WrapperExcel wrapper = new(pathToFile, listView, null)
                    {
                    };
                    wrapper.Exchange();
                }
                ),
                Dictionary<string, string> when value is Dictionary<string, string[]> dicStr => new Action(() =>
                {
                    DictionaryView listView = new(ExchangeOperation.Insert, sheetName, dicStr, null)
                    { };
                    WrapperExcel wrapper = new(pathToFile, listView, null)
                    {
                    };
                    wrapper.Exchange();
                }
                ),
                _ => new Action(()=>InsertListArray(value, pathToFile, sheetName, border)),
            };
            action.Invoke();
        }


        public static void GetFromExcel<TReturn>(out TReturn value, string pathToFile, string sheetName, Excel.Border border = null) where TReturn : new()
        {
            TReturn returnValue = new();
            if (returnValue is List<string> rL)
            {
                var ExcelExchange = new Excel.ListView(ExchangeOperation.Read, sheetName, rL, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, ExcelExchange, null) { };
                wrapper.Exchange();
                value = (TReturn)ExcelExchange.ExchangeValue;
                return;
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                var ExcelExchange = new Excel.DictionaryView(ExchangeOperation.Read, sheetName, rD, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, ExcelExchange, null) { };
                wrapper.Exchange();
                value = (TReturn)ExcelExchange.ExchangeValue;
                return;
            }
            else
            {
                value= GetListArrayChoise<TReturn>(returnValue,pathToFile, sheetName, border);
                return; 
            }
        }

        private static void InsertListArray<TValue>(TValue value, string pathToFile, string sheetName, Excel.Border border = null)
        {
            Action action = value switch
            {
                IList<string[]> when value is IList<string[]> listStr => new Action(() => InsertListArray<string>(listStr, pathToFile, sheetName, border)),
                IList<int[]> when value is IList<int[]> listInt => new Action(() => InsertListArray<int>(listInt, pathToFile, sheetName, border)),
                IList<double[]> when value is IList<double[]> listDbl => new Action(() => InsertListArray<Double>(listDbl, pathToFile, sheetName, border)),
                IList<bool[]> when value is IList<bool[]> listBool => new Action(() => InsertListArray<Boolean>(listBool, pathToFile, sheetName, border)),
                IList<DateTime[]> when value is IList<DateTime[]> listDateTime => new Action(() => InsertListArray<DateTime>(listDateTime, pathToFile, sheetName, border)),
                _ => default
            };
            action.Invoke();
        }



        private static ReturnValue GetListArrayChoise<ReturnValue>(dynamic value, string pathToFile, string sheetName, Excel.Border border = null) => value switch
        {
            IList<string[]> =>(ReturnValue)GetListArray<string>(pathToFile, sheetName, border),
            IList<int[]> => (ReturnValue)GetListArray<int>(pathToFile, sheetName, border),
            IList<double[]> => (ReturnValue)GetListArray<double>(pathToFile, sheetName, border),
            IList<bool[]> => (ReturnValue)GetListArray<bool>(pathToFile, sheetName, border),
            IList<DateTime[]> => (ReturnValue)GetListArray<DateTime>(pathToFile, sheetName, border),
            _ => default
        };

        private static IList<ReturnType[]> GetListArray<ReturnType>(string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
            var ExcelExchange = new MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
            WrapperExcel wrapper = new(pathToFile, ExcelExchange, null) { };
            wrapper.Exchange();
            return ExcelExchange.ExchangeValue;
        }

        private static void InsertListArray<TInsert>(IList<TInsert[]> value,string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
            var ExcelExchange = new MatrixViewGeneric<TInsert>(ExchangeOperation.Insert, sheetName, value, border, null);
            WrapperExcel wrapper = new(pathToFile, ExcelExchange, null) { };
            wrapper.Exchange();
        }
        /*
        public static void GetFromExcelListArray<ReturnType>(out IList<ReturnType[]> value, string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
                var ExcelExchange = new Excel.MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, ExcelExchange, null) { };
                wrapper.Exchange();
                value = ExcelExchange.ExchangeValue;
                return;
        }
        */
        public static void GetFromExcel(out DataFrame value, string pathToFile, string sheetName, Excel.Border border = null, Dictionary<int, Type> header = null, int[] rows = null)
        {
            var ExcelExchange = new Excel.DataFrameView(ExchangeOperation.Read, sheetName, null, border);
            if (rows != null)
            {
                ExcelExchange.DataHeader = new()
                {
                    Rows = rows
                };
            }
            else
            {
                ExcelExchange.DataHeader = new();
            }
            if (header != null)
            {
                ExcelExchange.DataHeader.CreateHeaderType(header);
            }
            Excel.WrapperExcel wrapper = new(pathToFile, ExcelExchange, null);
            wrapper.Exchange();
            value = ExcelExchange.ExchangeValue;
        }
        /// <summary>
        ///GetFromWord
        /// </summary>
        /// <param name="value"></param>
        /// <param name="pathToFile"></param>
        public static void GetFromWord(out List<Word.TableValue> value, string pathToFile)
        {
            var ExcelExchange = new Word.TableView(ExchangeOperation.Read);
            Word.WrapperWord wrapper = new(pathToFile, ExcelExchange, null);
            wrapper.Exchange();
            value = ExcelExchange.ExchangeValue;
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