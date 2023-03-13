©sashamsuper, 2020–2022
# WrapperNPOI
Wrapper for NPOI lib

Entering or introducing a list, dictionary or matrix into an Excel sheet
Example in testWrapper\MsTestWrapper\

The simplest use for receiving data

//DataFrame

const string path = "..//..//..//srcTest//dataframe.xlsx";
Simple.GetFromExcel(out DataFrame df, path, "Sheet1");
Debug.WriteLine(df);
//List<string>  
Simple.GetFromExcel(out List<string> ls, path, "Sheet1");  
Debug.WriteLine(String.Join("\n",ls));
//List<string[]>  
Simple.GetFromExcel(out List<string[]> lsm, path, "Sheet1");  
Debug.WriteLine(string.Join("\n", lsm.Select(x => string.Join("", x))));  
//Dictionary<string,string>  
Simple.GetFromExcel(out Dictionary<string, string[]> ld, path, "Sheet1");  
Debug.WriteLine(string.Join("\n", ld.Select(x=>$"Key:{x.Key}Value:{String.Join("",x.Value)}") ));
  

![example workflow](https://github.com/sashamsuper/WrapperNetPOI/actions/workflows/dotnet.yml/badge.svg)

