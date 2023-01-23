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
using NPOI.SS.UserModel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

            WrapperExcel wrapper = new(pathRec, rowsView, logger);
            wrapper.Exchange();
            rowsView.ExchangeValue.ToString();
        }

        public class SLogger
        {

            public static ILogger SimpleLogger()
            {
                var pathLog = ReturnTechFileName("Log", "Log");
                var logger = new LoggerConfiguration()
                     .MinimumLevel.Verbose() // 
                                             //
                     .WriteTo.File(pathLog) // 
                     .CreateLogger();
                return logger;
            }



            public static string ReturnTechFileName(string predict, string extension)
            {
                int i = 0;
                string rnd = "";
                var basePath = AppDomain.CurrentDomain.BaseDirectory ?? "";
                string dir = Path.Combine(basePath, predict);
                if (Directory.Exists(dir) == false)
                {
                    Directory.CreateDirectory(dir);
                }
                string path;
                do
                {
                    path = Path.Combine(dir, $"{predict}{DateTime.Now:yyMMddHHmmss}{rnd}.{extension}");
                    i += 1;
                    rnd = i.ToString();
                }
                while (File.Exists(path));
                return path;
            }

        }


        /// <summary>
        /// Defines the <see cref="Program" />.
        /// </summary>


        /// <summary>
        /// The Main.
        /// </summary>

        public static List<string[]> OneDay(List<string[]> list, int skipColumn = 0, int tableCount = 8, int dateColumn = 2, int rowColumn = 6)
        {
            List<string[]> tmpList = new();
            tmpList.AddRange(list);
            var date = tmpList[rowColumn].Skip(skipColumn).Take(tableCount).ElementAt(dateColumn);
            for (int i = 0; i < tmpList.Count; i++)
            {
                if (i != rowColumn)
                {
                    var listRow = tmpList[i].Skip(skipColumn).Take(tableCount).ToList();
                    listRow.Insert(0, date);
                    tmpList[i] = listRow.ToArray();
                }
            }
            return tmpList;
        }

        public static int[] FirstColumns(List<string[]> list)
        {
            List<int> intL = new();
            for (int i = 0; i < list[6].Length; i++)
            {
                if (list[6][i].Contains("Отчётные") || list[6][i].Contains("Отчетные"))
                {
                    intL.Add(i);
                }
            }
            return intL.ToArray();
        }
        public static void MainOLD()
        {

            //string pathToFile = @"C:\Users\Александр\source\repos\WrapperNPOI\WrapperNPOI\documents\book.xlsx";


            string path = "/storage/emulated/0/angle/dudu.xlsx";
            string[] df1 = { "23", "424", "33" };
            string[] df2 = { "23423", "234424", "2433" };
            string[] df3 = { "2a3423", "2d34424", "2sdf433" };

            List<string[]> outV = new(new[] { df1, df2, df3 });
            MatrixView matrix = new(ExchangeOperation.Insert, "Лист1", outV, null);
            WrapperExcel wrapper = new(path, matrix, SLogger.SimpleLogger());
            wrapper.Exchange();


            //SwapCellRange cells = new SwapCellRange(10, 10);
            Console.WriteLine(22);



        }
    }

}
