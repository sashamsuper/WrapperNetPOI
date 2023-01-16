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







