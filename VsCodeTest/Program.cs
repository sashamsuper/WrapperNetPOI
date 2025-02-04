using System.Xml.Serialization;
using System.Xml.Schema;
using WrapperNetPOI;
using System.IO;

namespace VsCodeTest;

class Program
{
    static void Main(string[] args)
    {
        var files = Directory.GetFiles(
            @"B:\Новая_папка\Отметки\Console\net6.0-windows\Attachments"
        );
        foreach (var file in files)
        {
            try
            {
                List<string[]> value = Simple.GetFromExcel<List<string[]>>(file, null, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(file);
            }
        }
    }
}
