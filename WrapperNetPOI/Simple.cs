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
        public static string[] GetSheetsNames(string pathToFile)
        {
            ListViewGeneric<string> listView = new(ExchangeOperation.Read, null, default);
            WrapperExcel wrapperExcel = new(pathToFile, listView);
            wrapperExcel.Exchange();
            return listView.SheetsNames;
        }

        

        public static void InsertToExcel<TInsert>(TInsert value, string pathToFile, string sheetName="Sheet1", Border border=null,Header header=null)
        {
            Action action = value switch
            {
                List<String> when value is List<string> listStr => new Action(() =>
                {
                    ListViewGeneric<string> listView = new(ExchangeOperation.Insert, sheetName, listStr, border)
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
                    DictionaryViewGeneric<string> listView = new(ExchangeOperation.Insert, sheetName, dicStr, border)
                    { };
                    WrapperExcel wrapper = new(pathToFile, listView, null)
                    {
                    };
                    wrapper.Exchange();
                }
                ),
                DataFrame when value is DataFrame dataFrame => new Action(() =>
                {
                    DataFrameView dataView = new(ExchangeOperation.Insert, sheetName, dataFrame, border,header)
                    { };
                    WrapperExcel wrapper = new(pathToFile, dataView, null)
                    {
                    };
                    wrapper.Exchange();
                }),
                _ => new Action(()=>InsertListArray(value, pathToFile, sheetName, border))
            };
            action.Invoke();
        }

        public static void UpdateToExcel<TUpdate>(TUpdate value, string pathToFile, string sheetName = "Sheet1", Border border = null)
        {
            Action action = value switch
            {
                List<String> when value is List<string> listStr => new Action(() =>
                {
                    ListViewGeneric<string> listView = new(ExchangeOperation.Update, sheetName, listStr, null)
                    {
                        ExchangeValue = listStr
                    };
                    WrapperExcel wrapper = new(pathToFile, listView, null)
                    {
                    };
                    wrapper.Exchange();
                }
                ),
                Dictionary<string, string> when value is Dictionary<string, string[]> dicStr => new Action(() =>
                {
                    DictionaryViewGeneric<string> listView = new(ExchangeOperation.Update, sheetName, dicStr, null)
                    { };
                    WrapperExcel wrapper = new(pathToFile, listView, null)
                    {
                    };
                    wrapper.Exchange();
                }
                ),
                _ => new Action(() => UpdateListArray(value, pathToFile, sheetName, border)),
            };
            action.Invoke();
        }

        public static void GetFromExcel<TReturn>(out TReturn value, string pathToFile, string sheetName, Excel.Border border = null) where TReturn : new()
        {
            TReturn returnValue = new();
            if (returnValue is List<string> rL)
            {
                var exchangeClass = new Excel.ListViewGeneric<string>(ExchangeOperation.Read, sheetName, rL, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (TReturn)exchangeClass.ExchangeValue;
                return;  
            }
            else if (returnValue is Dictionary<string, string[]> rD)
            {
                var exchangeClass = new Excel.DictionaryViewGeneric<string>(ExchangeOperation.Read, sheetName, rD, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = (TReturn)exchangeClass.ExchangeValue;
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

        private static void UpdateListArray<TValue>(TValue value, string pathToFile, string sheetName, Excel.Border border = null)
        {
            Action action = value switch
            {
                IList<string[]> when value is IList<string[]> listStr => new Action(() => UpdateListArray<string>(listStr, pathToFile, sheetName, border)),
                IList<int[]> when value is IList<int[]> listInt => new Action(() => UpdateListArray<int>(listInt, pathToFile, sheetName, border)),
                IList<double[]> when value is IList<double[]> listDbl => new Action(() => UpdateListArray<Double>(listDbl, pathToFile, sheetName, border)),
                IList<bool[]> when value is IList<bool[]> listBool => new Action(() => UpdateListArray<Boolean>(listBool, pathToFile, sheetName, border)),
                IList<DateTime[]> when value is IList<DateTime[]> listDateTime => new Action(() => UpdateListArray<DateTime>(listDateTime, pathToFile, sheetName, border)),
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
            var exchangeClass = new MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
            WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
            wrapper.Exchange();
            return exchangeClass.ExchangeValue;
        }

        private static void InsertListArray<TInsert>(IList<TInsert[]> value,string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
            var exchangeClass = new MatrixViewGeneric<TInsert>(ExchangeOperation.Insert, sheetName, value, border, null);
            WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
            wrapper.Exchange();
        }

        private static void UpdateListArray<TUpdate>(IList<TUpdate[]> value, string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
            var exchangeClass = new MatrixViewGeneric<TUpdate>(ExchangeOperation.Update, sheetName, value, border, null);
            WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
            wrapper.Exchange();
        }
        /*
        public static void GetFromExcelListArray<ReturnType>(out IList<ReturnType[]> value, string pathToFile, string sheetName, Excel.Border border = null) //where ReturnType : new()
        {
                var exchangeClass = new Excel.MatrixViewGeneric<ReturnType>(ExchangeOperation.Read, sheetName, null, border, null);
                Excel.WrapperExcel wrapper = new(pathToFile, exchangeClass, null) { };
                wrapper.Exchange();
                value = exchangeClass.ExchangeValue;
                return;
        }
        */
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