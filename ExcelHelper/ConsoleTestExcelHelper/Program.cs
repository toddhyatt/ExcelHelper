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
            using(ExcelHelper.ExcelApp app = new ExcelHelper.ExcelApp(@"c:\temp\test.txt"))
            {
                Console.WriteLine(app.DoesFileExist().ToString()); ;
                
            }
            Console.ReadLine();
        }
    }
}
