using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingUpdateWorksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            string templateFileName = @"C:\temp\AddDataTesting\InvoiceTemplate.xlsx";
            string newFile = @"C:\temp\AddDataTesting\InvoiceAdds.xlsx";
            ExcelHelper.ExcelApp app = new ExcelHelper.ExcelApp();
            app.OpenWorkBookTemplate(newFile, templateFileName);
            System.Data.DataTable dt = app.GetFirstOrDefault();
            app.InsertRow(newFile, "Sheet1", 6);
            app.UpdateRange(newFile, "Sheet1", 1,2, "newValue");
            app.SaveWorkBooks();
            
        }
    }
}
