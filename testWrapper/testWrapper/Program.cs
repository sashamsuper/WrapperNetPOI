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
        }
    }
}
