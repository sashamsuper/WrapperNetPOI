// <copyright file=Program.cs>
// Copyright (c) 2021 All Rights Reserved
// </copyright>
// <date>10.10.2021</date>

using System;
using System.Collections.Generic;
using System.Linq;

namespace WrapperNetPOI
{
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
            for (int i = 0; i < list[6].Length; i++)
            {
                if (list[6][i].Contains("Отчётные") || list[6][i].Contains("Отчетные"))
                {
                    intL.Add(i);
                }
            }
            return intL.ToArray();
        }




        public static void Main()
        {

            //string pathToFile = @"C:\Users\Александр\source\repos\WrapperNPOI\WrapperNPOI\documents\book.xlsx";


            string path = @"B:\dudu.xlsx";
            string[] df1 = { "23", "424", "33" };
            string[] df2 = { "23423", "234424", "2433" };
            string[] df3 = { "2a3423", "2d34424", "2sdf433" };
            List<string[]> outV = new(new[] { df1, df2, df3 });


            //SwapCellRange cells = new SwapCellRange(10, 10);
            Console.WriteLine(22);



            //if (File.Exists(path))
            //{
            //    File.Delete(path);
        }


        //Console.Write("Сколько строк сначала строки пропустить? ");
        //int.TryParse(Console.ReadLine(),out int propusk);
        //DirectoryInfo directoryInfo = new DirectoryInfo(Environment.CurrentDirectory);
        //var fileInfos=directoryInfo.GetFiles("*.xls*",SearchOption.TopDirectoryOnly);
        //Wrapper mainWrapper=new Wrapper("",ExchangeType.Update)





        //ICellRange<ICell> ddfdf = new Cell();


        //RangeView rangeView = new RangeView(ExchangeType.Get, "Лист1",);




        //wrapper.Exchange();


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

