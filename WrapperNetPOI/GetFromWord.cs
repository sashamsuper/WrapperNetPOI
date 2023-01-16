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
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace WrapperNetPOI
{
    public class WordExchange
    {

        public List<List<string[]>> Tables { set; get; } = new List<List<string[]>>();

        private void GetInformFromTable(IBody document)
        {
            List<string[]> rows = new();
            foreach (var table in document.Tables)
            {
                foreach (var row in table.Rows)
                {
                    string[] cells = default;
                    foreach (var cell in row.GetTableCells())
                    {
                        cells = row.GetTableCells().Select(x => x.GetText()).ToArray();
                    }
                    rows.Add(cells);
                }
                Tables.Add(rows);
            }
        }

        public void OpenFile(string filePath)
        {
            using FileStream file = new(filePath, FileMode.Open, FileAccess.Read);
            XWPFDocument document = new(file);
            GetInformFromTable(document);
        }
    }
}







