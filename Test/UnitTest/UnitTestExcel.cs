using System.Security.AccessControl;
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

using Microsoft.Data.Analysis;
using MsTestWrapper;
using NPOI.SS.UserModel;
using System;
using System.Diagnostics;
using System.Globalization;
using WrapperNetPOI;
using WrapperNetPOI.Excel;
using System.Reflection;

Console.WriteLine(22);
UnitTestExcel unitTestExcel = new();
unitTestExcel.SimpleGetFromExcel();

namespace MsTestWrapper

{
    [TestClass]
    public class UnitTestExcel
    {
        [TestMethod]
        public void ConvertTestDouble()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            RowsView exchangeClass = new(ExchangeOperation.Read);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            var value = exchangeClass.ExchangeValue;
            var row1 = value.Skip(1).First();
            var row2 = value.Skip(2).First();
            var row3 = value.Skip(3).First();
            ICell cell1 = row1.GetCell(0);
            ICell cell2 = row2.GetCell(0);
            ICell cell3 = row3.GetCell(0);
            ConvertType convertType = new();
            var d1 = convertType.GetValue<Double>(cell1);
            var d2 = convertType.GetValue<Double>(cell2);
            var d3 = convertType.GetValue<Double>(cell3);
            Assert.AreEqual(1, d1);
            Assert.AreEqual(2, d2);
            Assert.AreEqual(3, d3);
        }

        [TestMethod]
        public void ConvertTestInt()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            RowsView exchangeClass = new(ExchangeOperation.Read);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            var value = exchangeClass.ExchangeValue;
            var row1 = value.Skip(1).First();
            var row2 = value.Skip(2).First();
            var row3 = value.Skip(3).First();
            ICell cell1 = row1.GetCell(0);
            ICell cell2 = row2.GetCell(0);
            ICell cell3 = row3.GetCell(0);
            ConvertType convertType = new();
            var d1 = convertType.GetValue<Int32>(cell1);
            var d2 = convertType.GetValue<Int32>(cell2);
            var d3 = convertType.GetValue<Int32>(cell3);
            Assert.AreEqual(1, d1);
            Assert.AreEqual(2, d2);
            Assert.AreEqual(3, d3);
        }

        [TestMethod]
        public void ConvertTestString()
        {
            var path = "..//..//..//srcTest//dataframe.xlsx";
            RowsView exchangeClass = new(ExchangeOperation.Read);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            var value = exchangeClass.ExchangeValue[0];
            ICell cell = value.GetCell(0);
            ConvertType convertType = new();
            var str = convertType.GetValue<String>(cell);
            Assert.AreEqual("dron", str);
        }


        [TestMethod]
        public void ConvertToDictionaryTest()
        {
            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" },
                new[]{ "wer1","4wer","34wr" },
                new[]{ "wer1","5wer","34wr" },
            };
            var expected = new Dictionary<string, string[]>()
            {
                { "34", new [] { "2r3" }},
                { "1",new[]{"3we" }},
                { "wer1", new[]{"3wer","4wer","5wer"}}
            };
            var actual = Extension.ConvertToDictionary(listS);
            var expectedConv = expected.Select(x => (x.Key, String.Concat(x.Value))).ToList();
            var actualConv = actual.Select(x => (x.Key, String.Concat(x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
        }


        [TestMethod]
        public void DataFrameHeaderTest()
        {
            Header header = new()
            {
                Rows = new int[] { 0, 1 }
            };

            const string path = "..//..//..//srcTest//dataframe.xlsx";
            DataFrameView exchangeClass = new(ExchangeOperation.Read)
            {
                DataHeader = header
            };
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            Console.WriteLine(exchangeClass.DataHeader);
            var value = exchangeClass.ExchangeValue;
            string[] d = {"dron1", "header44", "sdds324", "asdrrg",
                    "asdg4",   "asd", "asd25",   "asd1" ,"asdaswer"};
            var value2 = exchangeClass.DataHeader.DataColumns.Select(x => x.Name).ToArray();
            CollectionAssert.AreEqual(d, value2);
        }

        [TestMethod]
        public void DataFrameHeaderTest2rowDiffLenght()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            DataFrameView exchangeClass = new(ExchangeOperation.Read, "Sheet2")
            {
                DataHeader = new()
                {
                    Rows = new int[] { 0, 1 }
                }
            };
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            Console.WriteLine(exchangeClass.DataHeader.DataColumns);
            var value = exchangeClass.ExchangeValue;
            string[] d = { "dron1", "header44",    "sdds324", "a11", "a244",
                "a3324",   "a41", "a744",    "asdas324" };
            var value2 = exchangeClass.DataHeader.DataColumns.Select(x => x.Name).ToArray();
            CollectionAssert.AreEqual(d, value2);
        }

        [TestMethod]
        public void DictionaryViewTestCreateInsert()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);
            Dictionary<string, string[]> dictSource = new()
            {
                { "1",new[]{"2","23","233" } },
                { "2",new[] { "2433", "24dfgd23", "dfg233" } },
                { "3",new[] { "34", "2dgd3", "2dgf33" } }
            };
            DictionaryView exchangeClass = new(ExchangeOperation.Insert, "List1", dictSource, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            exchangeClass = new(ExchangeOperation.Read, "List1", null, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var expectedConv = dictSource.Select(x => (x.Key, String.Concat(x.Value))).ToList();
            var actualConv = exchangeClass.ExchangeValue.Select(x => (x.Key, String.Concat(x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
            DeleteFile(path);
        }

        [TestMethod]
        public void DictionaryViewTestInsert()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            Dictionary<string, string[]> dictSource1 = new()
            {
                { "1",new[]{"2","23","233" } },
                { "2",new[] { "2433", "24dfgd23", "dfg233" } },
                { "3",new[] { "34", "2dgd3", "2dgf33" } }
            };
            DictionaryView exchangeClass = new(ExchangeOperation.Insert, "List1", dictSource1, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();

            Dictionary<string, string[]> dictSource2 = new()
            {
                { "3",new[]{"2342","23","23334" } },
                { "6",new[] { "2234433", "23244dfgd23", "��dfg233" } },
                { "7",new[] { "34234", "2342dgd3", "2dgf33��" } }
            };

            exchangeClass = new(ExchangeOperation.Insert, "List1", dictSource2, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();

            exchangeClass = new(ExchangeOperation.Read, "List1", null, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            foreach (var x in dictSource2)
            {
                if (dictSource1.ContainsKey(x.Key))
                {
                    var list = dictSource1[x.Key].ToList();
                    list.AddRange(x.Value);
                    dictSource1[x.Key] = list.ToArray();
                }
                else
                {
                    dictSource1.Add(x.Key, x.Value);
                }
            }
            var expectedConv = dictSource1.Select(x => (x.Key, string.Concat(x.Value))).ToList();
            var actualConv = exchangeClass.ExchangeValue.Select(x => (x.Key, String.Concat(x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
            DeleteFile(path);
        }

        [TestMethod]
        public void HandMaplexTestDouble()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            RowsView exchangeClass = new(ExchangeOperation.Read);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            var value = exchangeClass.ExchangeValue;
            var row3 = value.Skip(3).First();
            ICell cell3 = row3.GetCell(0);
            ConvertType convertType = new();
            Assert.AreEqual(3, convertType.GetValue<double>(cell3));
        }

        [TestMethod]
        public void ListViewTestCreateInsert()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            ListView exchangeClass = new(ExchangeOperation.Insert, "List1", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
            DeleteFile(path);
        }

        [TestMethod]
        public void ListViewTestCreateInsertFirstROWColumn()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };

            Border border = new()
            {
                FirstRow = 10,
                FirstColumn = 10
            };

            ListView exchangeClass = new(ExchangeOperation.Insert, "List1", listS, border, null)
            { };
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string> listGet = new();

            Border border2 = new()
            {
                FirstRow = 0,
                FirstColumn = 10
            };

            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, border2, null)
            { };
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            Assert.AreEqual(listS.Count + 10, exchangeClass.ExchangeValue.ToList().Count);
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.
                Where(x => !String.IsNullOrEmpty(x)).ToList());
            DeleteFile(path);
        }

        [TestMethod]
        public void ListViewTestInsert2Times()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            ListView exchangeClass = new(ExchangeOperation.Insert, "List1", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            listS.AddRange(listS);
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
            DeleteFile(path);
        }

        [TestMethod]
        public void ListViewTestUpdate()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);
            ListView pusto = new(ExchangeOperation.Insert, "List1", null, null);
            WrapperExcel wrapper = new(path, pusto, null);
            wrapper.Exchange();

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            ListView exchangeClass = new(ExchangeOperation.Update, "List1", listS, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();

            listS = new()
            {
                "54",
                "245",
                "345"
            };
            exchangeClass = new(ExchangeOperation.Update, "List1", listS, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            exchangeClass = new(ExchangeOperation.Read, "List1", null, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.ToList());
            DeleteFile(path);
        }

        [TestMethod]
        public void MatrixViewTestCreateInsert()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Insert, "List1455", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string[]> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var expected = listS.ConvertAll(x => String.Concat(x));
            var actual = exchangeClass.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        //[TestMethod]
        public void MatrixViewTest()
        {
            const string path = @"B:\Новая папка\Отметки водохранилищ 01.09.2021-22.09.2021\20220104_maketDPVBY_2022-01-03.xls";
            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Read, null, null, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var expected = listS.ConvertAll(x => String.Concat(x));
            var actual = exchangeClass.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }





        [TestMethod]
        public void MatrixViewGenericTest()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Insert, "List1455", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            //exchangeClass.Dispose();

            MatrixViewGeneric<string> exchangeClass2 = new(ExchangeOperation.Read, null, null, null);
            wrapper = new(path, exchangeClass2, null);
            wrapper.Exchange();
            var expected = listS.ConvertAll(x => String.Concat(x));
            var actual = exchangeClass2.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        [TestMethod]
        public void MatrixViewGenericIntTest()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Insert, "List1455", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            //exchangeClass.Dispose();

            MatrixViewGeneric<int> exchangeClass2 = new(ExchangeOperation.Read, null, null, null);
            wrapper = new(path, exchangeClass2, null);
            wrapper.Exchange();

            List<string> expected = new();
            foreach (var x in listS)
            {
                var row = new List<string>();
                foreach (var y in x)
                {
                    int.TryParse(y, out var intValue);
                    var value=intValue.ToString();
                    row.Add(value);
                }
                expected.Add(String.Concat(row.ToArray()));
            }
            var actual = exchangeClass2.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        [TestMethod]
        public void MatrixViewTestInsert2Times()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Insert, "List1", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string[]> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            listS.AddRange(listS);
            var expected = listS.ConvertAll(x => String.Concat(x));
            var actual = exchangeClass.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        [TestMethod]
        public void MatrixViewTestUpdate()
        {
            const string path = "..//..//..//srcTest//listView.xlsx";
            DeleteFile(path);
            ListView pusto = new(ExchangeOperation.Insert, "List1", null, null);
            WrapperExcel wrapper = new(path, pusto, null);
            wrapper.Exchange();

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Update, "List1", listS, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();

            listS = new()
            {
                new []{ "1","2","3" },
                new[]{ "1","3we","34" },
                new[]{ "1","3r","3r" }
            };
            exchangeClass = new(ExchangeOperation.Update, "List1", listS, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string[]> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null);
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var expected = listS.ConvertAll(x => string.Concat(x));
            var actual = exchangeClass.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        [TestMethod]
        public void ReturnProgressTest()
        {
            Assert.AreEqual(10, ExchangeClass<int>.ReturnProgress(10, 100));
            Assert.AreEqual(50, ExchangeClass<int>.ReturnProgress(50, 100));
            Assert.AreEqual(25, ExchangeClass<int>.ReturnProgress(25, 100));
        }

        [TestMethod]
        public void TestCopyExcelToExcel()
        {
            var path = "..//..//..//srcTest//listViewXLSX.xlsx";
            DeleteFile(path);

            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            MatrixView exchangeClass = new(ExchangeOperation.Insert, "List1455", listS, null);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();

            RowsView rowsView = new(ExchangeOperation.Insert, "List1455", null, null)
            {
                PathSource = path
            };
            const string path2 = "..//..//..//srcTest//listViewXLSX2.xlsx";
            if (File.Exists(path2))
            {
                File.Delete(path2);
            }
            wrapper = new(path2, rowsView, null);
            wrapper.Exchange();

            MatrixView matrix = new(ExchangeOperation.Read, null, null, null);
            wrapper = new(path2, matrix, null);
            wrapper.Exchange();

            var expected = listS.ConvertAll(x => String.Concat(x));
            var actual = matrix.ExchangeValue.Select(x => String.Concat(x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
            DeleteFile(path);
        }

        //[TestMethod]
        public void TestManyReadXLSX()
        {
            IProgress<int> progress = new Progress<int>(s => Debug.WriteLine(s));

            const string path = "..//..//..//srcTest//500000_Records_Data.xlsx";
            MatrixView exchangeClass = new(ExchangeOperation.Read, "List1", null, null, progress);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var dd = exchangeClass.ExchangeValue;
            Assert.AreEqual(400000, dd.Count);
        }

        [TestMethod]
        public void TestReadXLS()
        {
            IProgress<int> progress = new Progress<int>(s => Debug.WriteLine(s));

            const string path = "..//..//..//srcTest//listView3.xls";
            MatrixView exchangeClass = new(ExchangeOperation.Read, "List1", null, null, progress);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var dd = exchangeClass.ExchangeValue;
            //Assert.AreEqual(5,dd.Count);
        }

        [TestMethod]
        public void SimpleGetFromExcelBorder()
        {
            //DataFrame
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            Simple.GetFromExcel(out DataFrame df, path, "Sheet4",
                    new Border
                    {
                        FirstColumn = 5,
                        FirstRow = 5,
                        LastColumn = 10,
                        LastRow = 10
                    });
            Debug.WriteLine(df);
        }

        [TestMethod]
        public void SimpleGetFromExcel()
        {
            //DataFrame
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            Simple.GetFromExcel(out DataFrame df, path, "Sheet1",
                    new Border
                    {
                        FirstColumn = 5,
                        FirstRow = 5,
                        LastColumn = 10,
                        LastRow = 10
                    });
            Debug.WriteLine(df);
            //List<string>
            Simple.GetFromExcel(out List<string> ls, path, "Sheet1");
            Debug.WriteLine(String.Join("\n", ls));
            //List<string[]>
            Simple.GetFromExcel(out List<string[]> lsm, path, "Sheet1");
            Debug.WriteLine(string.Join("\n", lsm.Select(x => string.Concat(x))));
            //Dictionary<string,string>
            Simple.GetFromExcel(out Dictionary<string, string[]> ld, path, "Sheet1");
            Debug.WriteLine(string.Join("\n", ld.Select(x => $"Key:{x.Key}Value:{String.Concat(x.Value)}")));
        }


        [TestMethod]
        public void SimpleGetFromExcelWithGeneric()
        {
            //DataFrame
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            Simple.GetFromExcel(out DataFrame df, path, "Sheet1",
                    new Border
                    {
                        FirstColumn = 5,
                        FirstRow = 5,
                        LastColumn = 10,
                        LastRow = 10
                    });
            Debug.WriteLine(df);
            //List<string>
            Simple.GetFromExcel<List<string>>(out List<string> ls, path, "Sheet1");
            Debug.WriteLine(String.Join("\n", ls));
            //List<string[]>
            Simple.GetFromExcel<List<string[]>>(out List<string[]> lsm, path, "Sheet1");
            Debug.WriteLine(string.Join("\n", lsm.Select(x => string.Concat(x))));
            //Dictionary<string,string>
            Simple.GetFromExcel<Dictionary<string, string[]>>(out Dictionary<string, string[]> ld, path, "Sheet1");
            Debug.WriteLine(string.Join("\n", ld.Select(x => $"Key:{x.Key}Value:{String.Concat(x.Value)}")));
        }

        [TestMethod]
        public void DataFrameTestValue()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            Simple.GetFromExcel(out DataFrame df, path, "Sheet3");

            Dictionary<int, Type> header = new()
            {   { 0, typeof(String) } ,
                { 1, typeof(String) },
                { 2, typeof(String) }
            };
            var col1 = new StringDataFrameColumn("col1", new string[] { "1", "3", "6" });
            var col2 = new StringDataFrameColumn("col2",
                new String[] { "2", "4", "7" });
            var col3 = new StringDataFrameColumn("col3",
                new String[] { "3,1", "5,1", "8,1" });
            var sample = new DataFrame(col1, col2, col3).Rows.Select(x => x.ToString()).ToList();
            CollectionAssert.AreEqual(sample, df.Rows.Select(x => x.ToString()).ToList());
        }

        [TestMethod]
        public void DataFrameTestValueSimple()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            DataFrameView exchangeClass = new(ExchangeOperation.Read, "Sheet3")
            {
                DataHeader = new()
                {
                    Rows = new int[] { 0 }
                }
            };

            Dictionary<int, Type> header = new()
            {   { 0, typeof(String) } ,
                { 1, typeof(DateTime) },
                { 2, typeof(Double) }
            };
            exchangeClass.DataHeader.CreateHeaderType(header);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            Debug.WriteLine(exchangeClass.ExchangeValue);
            var value = exchangeClass.ExchangeValue;
            var col1 = new StringDataFrameColumn("col1", new string[] { "1", "3", "6" });



            DateTime.TryParse("4", CultureInfo.CurrentCulture, DateTimeStyles.AssumeUniversal,
            out var outValue2);



            var col2 = new DateTimeDataFrameColumn("col2",
                new DateTime[] { DateTime.FromOADate(2), outValue2, DateTime.FromOADate(7) });
            var col3 = new DoubleDataFrameColumn("col3",
                new Double[] { 3.1, 5.1, 8.1 });
            var sample = new DataFrame(col1, col2, col3).Rows.Select(x => x.ToString()).ToList();
            var value2 = exchangeClass.ExchangeValue.Rows.Select(x => x.ToString()).ToList();
            CollectionAssert.AreEqual(sample, value2);
        }






        [TestMethod]
        public void DataFrameIntegerSimple()
        {
            const string path = "..//..//..//srcTest//dataframe.xlsx";
            DataFrameView exchangeClass = new(ExchangeOperation.Read, "Sheet3")
            {
                DataHeader = new()
                {
                    Rows = new int[] { 0 }
                }
            };

            Dictionary<int, Type> header = new()
            {   { 0, typeof(int) } ,
                { 1, typeof(DateTime) },
                { 2, typeof(Double) }
            };
            exchangeClass.DataHeader.CreateHeaderType(header);
            WrapperExcel wrapper = new(path, exchangeClass);
            wrapper.Exchange();
            Debug.WriteLine(exchangeClass.ExchangeValue);
            var value = exchangeClass.ExchangeValue;
            var col1 = new Int32DataFrameColumn("col1", new int[] { 1, 3, 6 });

            DateTime.TryParse("4", CultureInfo.CurrentCulture, DateTimeStyles.AssumeUniversal,
            out var outValue);

            var col2 = new DateTimeDataFrameColumn("col2",
                new DateTime[] { DateTime.FromOADate(2), outValue, DateTime.FromOADate(7) });
            var col3 = new DoubleDataFrameColumn("col3",
                new Double[] { 3.1, 5.1, 8.1 });
            var sample = new DataFrame(col1, col2, col3).Rows.Select(x => x.ToString()).ToList();
            var value2 = exchangeClass.ExchangeValue.Rows.Select(x => x.ToString()).ToList();
            CollectionAssert.AreEqual(sample, value2);
        }

        protected static void DeleteFile(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }


        public static Array GetArrayFromList(dynamic value) => value switch
        {
            IList<string[]> => Array.Empty<string>(),
            IList<int[]> => Array.Empty<int>(),
            IList<double[]> => Array.Empty<double>(),
            IList<bool[]> => Array.Empty<bool>(),
            _ => Array.Empty<dynamic>()
        };

        public static void ReflectionView<T>(T value)
        {
            Type type = value.GetType();
            string fullName = type.FullName;
            Debug.WriteLine($"Generic-{ type.IsGenericType}");
            Debug.WriteLine($"Full Name-{type.FullName}");
            Debug.WriteLine(type.Name);
        } 
        
        
        [TestMethod]
        public void GetReflection()
        {
            List<string[]> DDD= new();
            var d= GetArrayFromList(DDD);
            ReflectionView(DDD);
        }
    }
}