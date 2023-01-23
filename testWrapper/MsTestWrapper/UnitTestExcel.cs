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
using System.Diagnostics;
using WrapperNetPOI;

namespace MsTestWrapper

{
    [TestClass]
    public class UnitTestExcel
    {
        [TestMethod]
        public void ReturnProgressTest()
        {
            Assert.AreEqual(10, ExchangeClass<int>.ReturnProgress(10, 100));
            Assert.AreEqual(50, ExchangeClass<int>.ReturnProgress(50, 100));
            Assert.AreEqual(25, ExchangeClass<int>.ReturnProgress(25, 100));
        }

        [TestMethod]
        public void ListViewTestCreateInsert()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
        }

        [TestMethod]
        public void ListViewTestCreateInsertFirstROWColumn()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            List<string> listS = new()
            {
                "1",
                "2",
                "3"
            };
            ListView exchangeClass = new(ExchangeOperation.Insert, "List1", listS, null)
            {
                FirstViewedRow = 10,
                FirstViewedColumn = 10
            };
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            List<string> listGet = new();
            exchangeClass = new(ExchangeOperation.Read, "List1", listGet, null)
            {
                FirstViewedRow = 0,
                FirstViewedColumn = 0
            };
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            Assert.AreEqual(listS.Count + 10, exchangeClass.ExchangeValue.ToList().Count);
            CollectionAssert.AreEqual(listS, exchangeClass.ExchangeValue.
                Where(x => String.IsNullOrEmpty(x) == false).ToList());
        }



        [TestMethod]
        public void ListViewTestUpdate()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }
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
        }

        [TestMethod]
        public void ListViewTestInsert2Times()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
        }


        [TestMethod]
        public void MatrixViewTestCreateInsert()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
            var expected = listS.Select(x => String.Join("", x)).ToList();
            var actual = exchangeClass.ExchangeValue.Select(x => String.Join("", x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void MatrixViewTestInsert2Times()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
            var expected = listS.Select(x => String.Join("", x)).ToList();
            var actual = exchangeClass.ExchangeValue.Select(x => String.Join("", x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void MatrixViewTestUpdate()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }
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
            var expected = listS.Select(x => String.Join("", x)).ToList();
            var actual = exchangeClass.ExchangeValue.Select(x => String.Join("", x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void DictionaryViewTestCreateInsert()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
            var expectedConv = dictSource.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            var actualConv = exchangeClass.ExchangeValue.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
        }

        [TestMethod]
        public void DictionaryViewTestInsert()
        {
            var path = "..//..//..//srcTest//listView.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
            var expectedConv = dictSource1.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            var actualConv = exchangeClass.ExchangeValue.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
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
            var expectedConv = expected.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            var actualConv = actual.Select(x => (x.Key, String.Join("", x.Value))).ToList();
            CollectionAssert.AreEqual(expectedConv, actualConv);
        }

        [TestMethod]
        public void TestReadXLS()
        {
            IProgress<int> progress = new Progress<int>(s => Debug.WriteLine(s));

            var path = "..//..//..//srcTest//listView2.xls";
            MatrixView exchangeClass = new(ExchangeOperation.Read, "List1", null, progress);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var dd = exchangeClass.ExchangeValue;
            //Assert.AreEqual(5,dd.Count);
        }


        [TestMethod]
        public void TestManyReadXLSX()
        {
            IProgress<int> progress = new Progress<int>(s => Debug.WriteLine(s));

            var path = "..//..//..//srcTest//500000_Records_Data.xlsx";
            MatrixView exchangeClass = new(ExchangeOperation.Read, "List1", null, progress);
            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var dd = exchangeClass.ExchangeValue;
            Assert.AreEqual(400000, dd.Count);
        }

        [TestMethod]
        public void TestCopyExcelToExcel()
        {
            var path = "..//..//..//srcTest//listViewXLSX.xlsx";
            if (File.Exists(path))
            {
                File.Delete(path);
            }

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
            var path2 = "..//..//..//srcTest//listViewXLSX2.xlsx";
            if (File.Exists(path2))
            {
                File.Delete(path2);
            }
            wrapper = new(path2, rowsView, null);
            wrapper.Exchange();

            MatrixView matrix = new(ExchangeOperation.Read, null, null, null);
            wrapper = new(path2, matrix, null);
            wrapper.Exchange();

            var expected = listS.Select(x => String.Join("", x)).ToList();
            var actual = matrix.ExchangeValue.Select(x => String.Join("", x)).ToList();
            CollectionAssert.AreEqual(expected, actual);
        }
    }
}