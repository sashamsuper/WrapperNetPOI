// <copyright file=Program.cs>
// Copyright (c) 2021 All Rights Reserved
// </copyright>
// <date>10.10.2021</date>

namespace WrappperNPOI
{
    using System;
    using System.Collections.Generic;
    using System.Linq;


    /// <summary>
    /// Defines the <see cref="Program" />.
    /// </summary>
    public class Program
    {
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
            for (int i = 0; i<list[6].Length; i++)
            {
                if (list[6][i].Contains("Отчётные")|| list[6][i].Contains("Отчетные"))
                {
                    intL.Add(i);
                }    
            }
            return intL.ToArray();
        }




        public static void Main()
        {

            //string pathToFile = @"C:\Users\Александр\source\repos\WrapperNPOI\WrapperNPOI\documents\book.xlsx";

            string path = @"/storage/F07E-171B/Downloads/TAX_REPORT.xlsx";
            //string pathToFile = @"B:\document.xlsx";
            //string pathToFile = @"D:\tmp\Печорская\21.01.22.docx";
            var d=ExcelExchange.GetFromExcel<List<string[]>>(path,"Лист1");
            foreach (var x in d)
            {
            	Console.WriteLine(String.Join(";",x));
            }
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
