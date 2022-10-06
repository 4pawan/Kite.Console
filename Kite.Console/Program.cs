using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Zerodha.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Directory.GetCurrentDirectory();
            Console.WriteLine(path + "\n\n");
            Console.WriteLine("Use Monthly JSON to generate report...");
            Console.WriteLine("Use 5 min candle chart generate intraday report...");
            Console.WriteLine("Enter Choice 1 to generate daily chart excel...");
            Console.WriteLine("Enter Choice 2 to fill intraday details...This can be generated after option 1...");
            var input = Console.ReadKey();
            Console.WriteLine("Generating excel now..");
            Excelhelper.ExportToExcel(input.Key.ToString());
            Console.WriteLine("------Press any key to exit! -----------------");
            Console.ReadKey();
        }
    }
}
