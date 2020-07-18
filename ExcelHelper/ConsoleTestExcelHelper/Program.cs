using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTestExcelHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            using(ExcelHelper.ExcelApp app = new ExcelHelper.ExcelApp())
            {
                string templateName = @"C:\temp\ConsoleTestExcelHelper\Template.xltx";
                //string templateName = "Template.xltx";

                app.OpenWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile1.xlsx");
                app.OpenWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile2.xlsx");
                app.SaveWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile1.xlsx");
                app.SaveWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile2.xlsx");
                app.OpenWorkBookTemplate(@"C:\temp\ConsoleTestExcelHelper\TestFile3t.xlsx",templateName);
                app.SaveWorkBooks();

                Console.WriteLine("going to dispose now...");
            }
            Console.WriteLine("Press return key to end.");
            Console.ReadLine();
        
       }
    }
}
