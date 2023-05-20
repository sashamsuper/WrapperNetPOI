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
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WrapperNetPOI;
using WrapperNetPOI.Excel;

namespace UnitTest
{
    [TestClass]
    public class UnitExcelInsert
    {
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
            Simple.InsertToExcel(listS, path, "SheetNew",null);
            Simple.GetFromExcel(out List<string[]> ls, path, "SheetNew");
            Debug.WriteLine(String.Join("\n", ls));
            CollectionAssert.AreEqual(listS.SelectMany(x=>x).ToArray(),ls.SelectMany(x=>x).ToArray());
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
            Simple.InsertToExcel(listS, path, "SheetNew", new Border(firstColumn:5,firstRow:5));
            Simple.GetFromExcel(out List<int[]> ls, path, "SheetNew", new Border(firstColumn: 5, firstRow: 5));
            Debug.WriteLine(String.Join("\n", ls));
            CollectionAssert.AreEqual(listS.SelectMany(x => x).ToArray(), ls.SelectMany(x => x).ToArray());
        }
    }
}
