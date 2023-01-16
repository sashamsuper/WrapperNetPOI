using NPOI.SS.UserModel;
using Serilog;
using System.Collections.Generic;
using WrapperNetPOI;

namespace testWrapper
{
    class SwapCell
    {
        public int Width { get; }
        public int Height { get; }
        public (int, int) TopLeftCell { get; set; } = (0, 0);
        public ICell[][] Cells { get; }
        public string Path { get; }
        public ICell GetCell(int row, int column)
        {
            (int firstRow, int firstCol) = TopLeftCell;
            return Cells[firstRow + row][firstCol + column];
        }

        public SwapCell(string path, int height, int width)
        {
            Path = path;
            Height = height;
            Width = width;
        }
    }

    internal class Program
    {


        public static void Main()
        {

            //string pathToFile = @"C:\Users\Александр\source\repos\WrapperNPOI\WrapperNPOI\documents\book.xlsx";

            string pathSource = @"B:\dudu.xls";
            string pathRec = @"B:\dudu2.xlsx";
            var pathLog = Wrapper.ReturnTechFileName("Log", "log");
            ILogger logger = new LoggerConfiguration()
             .MinimumLevel.Verbose() // ставим минимальный уровень в Verbose для теста, по умолчанию стоит Information 
                                     //.WriteTo.Console()  // выводим данные на консоль
             .WriteTo.File(pathLog) // а также пишем лог файл, разбивая его по дате
             .CreateLogger();
            RowsView rowsView = new(ExchangeOperation.Update, "Лист1", new List<IRow>(), null)
            {
                LastViewedRow = 90,
                PathSource = pathSource,
            };

            Wrapper wrapper = new Wrapper(pathRec, rowsView, logger);
            wrapper.Exchange();
            rowsView.ExchangeValue.ToString();





            /*
            string[] df1 = { "23", "424", "33" };
            string[] df2 = { "23423", "234424", "2433" };
            string[] df3 = { "2a3423", "2d34424", "2sdf433" };
            string[] df4 = { "2a3423", "2d34424", "2sdf433" };
            string[] df5= { "2a3423", "2d34424", "2sdf433" };
            string[] df6 = { "2a3423", "2d34424", "2sdf433" };
            string[] df7 = { "2a3423", "2d34424", "2sdf433" };
            string[] df87 = { "2a3423", "2d34424", "2sdf433" };
            string[] df456 = { "2a3423", "2d34424", "2sdf433" };
            List<string[]> outV = new()
            {
                df1,
                df2,
                df3,
                df4,
                df5,
                df6,
                df7,
                df87,
                df456
            };

            if (File.Exists(path))
            { 
                File.Delete(path);
            }
*/




            /*
                    MatrixView mView = new(ExchangeType.Add, "Лист1", outV)
                        {
                            progress = new Progress<double>((x)=>Console.WriteLine(x))
                        };
                        Wrapper wrapper = new(path, mView);



                        wrapper.Exchange();

                    */

            //WrapperNpoi

            //ExcelExchange.AddToExcel(path, "Лист1", outV);

            //string pathToFile = @"B:\document.xlsx";
            //string pathToFile = @"D:\tmp\Печорская\21.01.22.docx";
            /*
            var d=ExcelExchange.GetFromExcel<List<string[]>>(path,"Лист1");
            foreach (var x in d)
            {
            	Console.WriteLine(String.Join(";",x));
            }
            */
            //GetFromWord getFromWord = new GetFromWord();
            //getFromWord.OpenFile(pathToFile);
            //Console.WriteLine(getFromWord.Tables);
            /*
            var paths = Directory.GetFiles(pathToFiles);
            List<string[]> list = new List<string[]>();
            foreach (var path in paths)
            {
                List<string[]> tmpList = ExchangeExcel.GetFromExcel<List<string[]>>(path, "Лист1");
                int[] firstColumns = FirstColumns(tmpList);
                List<string[]> partTmpList = new List<string[]>();
                foreach (int col in firstColumns)
                {
                    partTmpList=OneDay(tmpList,skipColumn:col);
                    tmpList.AddRange(partTmpList);
                }
                list.AddRange(tmpList);
            }
            list = list.Where(x => String.IsNullOrWhiteSpace(x.ElementAtOrDefault(3)) == false).ToList();
            var newPath = @"B:\TEMP\new.xlsx";
            ExchangeExcel.AddToExcel(newPath, "Лист1", list);
            Console.WriteLine(list);*/
        }
    }
}
