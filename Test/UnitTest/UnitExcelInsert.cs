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
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WrapperNetPOI;
using WrapperNetPOI.Excel;
using NPOI.SS.UserModel;

namespace UnitTest
{
    [TestClass]
    public class UnitExcelInsert
    {

        protected static void DeleteFile(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }


        [TestMethod]
        public void DataFrameTestInsert()
        {
            var path = Path.GetFullPath("..//..//..//srcTest//dataFrameInsert.xlsx");
            DeleteFile(path);

            var col1 = new StringDataFrameColumn("col1", new string[] { "a1", "a3", "6" });
            var col2 = new Int32DataFrameColumn("col2",
                new int[] { 2, 4, 7 });
            var col3 = new DoubleDataFrameColumn("col3",
                new Double[] { 3.1, 5.1, 8.1 });
            var sample = new DataFrame(col1, col2, col3);
            DataFrameView exchangeClass = new(ExchangeOperation.Insert, "List1", sample, null);
            Header header = new(Array.Empty<int>(), new Dictionary<int, Type>
            {
                {0,typeof(string)},
                {1,typeof(int)},
                {2,typeof(double)}
            });

            WrapperExcel wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            exchangeClass = new(ExchangeOperation.Read, "List1", null, null, header);
            //exchangeClass.DataHeader.Rows= Array.Empty<int>();
            wrapper = new(path, exchangeClass, null);
            wrapper.Exchange();
            var s1 = sample.Rows.SelectMany(x => x).ToList();
            var s2 = exchangeClass.ExchangeValue.Rows.SelectMany(x => x).ToList();
            CollectionAssert.AreEqual(s1, s2);
            //DeleteFile(path);
        }

        [TestMethod]
        public void SimpleGetFromExcelString()
        {
            //DataFrame
            const string path = "..//..//..//srcTest//simpleGeneric.xlsx";
            File.Delete(path);
            List<string[]> listS = new()
            {
                new []{ "34","2r3","34" },
                new[]{ "1","3we","34" },
                new[]{ "wer1","3wer","34wr" }
            };
            Simple.InsertToExcel(listS, path, "SheetNew", null);
            Simple.GetFromExcel(out List<string[]> ls, path, "SheetNew");
            Debug.WriteLine(String.Join("\n", ls));
            CollectionAssert.AreEqual(listS.SelectMany(x => x).ToArray(), ls.SelectMany(x => x).ToArray());
        }

        [TestMethod]
        public void SimpleGetFromExcelInt()
        {
            //DataFrame
            const string path = "..//..//..//srcTest//simpleGeneric.xlsx";
            File.Delete(path);
            List<int[]> listS = new()
            {
                new []{ 34,3,34 },
                new[]{ 1,55,34354 },
                new[]{ 1,3,4 }
            };
            Simple.InsertToExcel(listS, path, "SheetNew", new Border(firstColumn: 5, firstRow: 5));
            Simple.GetFromExcel(out List<int[]> ls, path, "SheetNew", new Border(firstColumn: 5, firstRow: 5));
            Debug.WriteLine(String.Join("\n", ls));
            CollectionAssert.AreEqual(listS.SelectMany(x => x).ToArray(), ls.SelectMany(x => x).ToArray());
        }

            }





}
