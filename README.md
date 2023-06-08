©sashamsuper, 2020–2023
# WrapperNetPOI
Wrapper for NPOI lib

Entering or introducing a list, dictionary or matrix into an Excel sheet
Example in testWrapper\MsTestWrapper\

The simplest use for receiving data

    using WrapperNetPOI;
    
    //DataFrame
    const string path = "..//..//..//srcTest//dataframe.xlsx";  
    Simple.GetFromExcel(out DataFrame df, path, "Sheet1");  
    Console.WriteLine(df);  

    //List<string>  
    Simple.GetFromExcel(out List<string> ls, path, "Sheet1");  
    Debug.WriteLine(String.Join("\n",ls));  

    //Dictionary<string,string>  
    Simple.GetFromExcel(out Dictionary<string, string[]> ld, path, "Sheet1");  
    Debug.WriteLine(string.Join("\n", ld.Select(x=>$"Key:{x.Key}Value:{String.Join("",x.Value)}") ));

    //List<string[]>  (work with string, int, double, DateTime, bool)
    Simple.GetFromExcel(out List<string[]> lsm, path, "Sheet1");  
    Debug.WriteLine(string.Join("\n", lsm.Select(x => string.Join("", x))));

The simplest use for insert data (work with string, int, double, DateTime, bool)

    using WrapperNetPOI;
    
    const string path = "..//..//..//srcTest//simpleGeneric.xlsx";  
    File.Delete(path);  
    List<string[]> listS = new()
        {
            new []{ "34","2r3","34" },
            new[]{ "1","3we","34" },
            new[]{ "wer1","3wer","34wr" }
        };
    Simple.InsertToExcel(listS, path, "SheetNew",null);

https://github.com/sashamsuper/WrapperNetPOI  
  

![example workflow](https://github.com/sashamsuper/WrapperNetPOI/actions/workflows/dotnet.yml/badge.svg)

