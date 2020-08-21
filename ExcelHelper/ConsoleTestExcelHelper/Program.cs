using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTestExcelHelper
{
    class Program
    {
        static void Main(string[] args)
        {
            TestModel();
            using (ExcelHelper.ExcelApp app = new ExcelHelper.ExcelApp())
            {
                using (ExcelHelper.ExcelApp app2 = new ExcelHelper.ExcelApp())
                {
                    string templateName = @"C:\temp\ConsoleTestExcelHelper\Template.xltx";
                    //string templateName = "Template.xltx";
                    app.OpenWorkBook(@"C:\temp\Event #5 - FULL EVENT REPORT.xls");
                    app2.OpenWorkBook(@"C:\temp\EquipmentRate.xlsx");

                    //app.OpenWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile1.xlsx");
                    //app.OpenWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile2.xlsx");
                    //app.SaveWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile1.xlsx");
                    //app.SaveWorkBook(@"C:\temp\ConsoleTestExcelHelper\TestFile2.xlsx");
                    //app.OpenWorkBookTemplate(@"C:\temp\ConsoleTestExcelHelper\TestFile3t.xlsx",templateName);
                    //app.SaveWorkBooks();
                    //var results = from myRow in myDataTable.AsEnumerable()
                    //              where myRow.Field<int>("RowNo") == 1
                    //              select myRow;
                    var eventDescr = app.GetFirstOrDefault().AsEnumerable().FirstOrDefault()[0].ToString();
                    System.Diagnostics.Trace.WriteLine(eventDescr.ToString());
                    var columnNames = from myRow in app.GetFirstOrDefault().Rows[5].ItemArray
                                      select myRow;
                    string commaDelimitedColumnNames = "";
                    foreach (var columnName in columnNames)
                    {
                        commaDelimitedColumnNames += columnName.ToString() + ",";
                        System.Diagnostics.Trace.WriteLine(columnName.ToString());
                    }
                    commaDelimitedColumnNames += "LaborRate,EquipmentRate";
                    var data = from myRow in app.GetFirstOrDefault().AsEnumerable().Skip(8)
                               where myRow[3] != null
                               select myRow;
                    System.Diagnostics.Trace.WriteLine(data.Count().ToString());
                    var subs = from subrow in data
                               where subrow[3].ToString().Trim().StartsWith("SUB")
                               orderby subrow[3].ToString()
                               select subrow;
                    System.Diagnostics.Trace.WriteLine(subs.Count().ToString());
                    string CurrentSub = "";
                    int LaborRate = 0;
                    int EquipmentRate = 0;
                    StringBuilder sb = new StringBuilder();

                    var rateColumnNames = from myRow in app2.GetFirstOrDefault().Rows[0].ItemArray
                                          select myRow;
                    int iLaborColumnNumber = -1;
                    int currentColumn = 0;
                    foreach (var col in rateColumnNames)
                    {
                        if (col.ToString() == "Labor")
                        {
                            iLaborColumnNumber = currentColumn;
                            break;
                        }
                        currentColumn += 1;

                    }

                    foreach (var DataItem in subs)
                    {
                        int ColCount = 0;
                        if (CurrentSub != DataItem[3].ToString())
                        {
                            if (CurrentSub != "")
                                WriteSbToCSV(eventDescr, CurrentSub, sb);
                            CurrentSub = DataItem[3].ToString();
                            sb = new StringBuilder();
                            sb.AppendLine(commaDelimitedColumnNames);
                            LaborRate = 0;
                            EquipmentRate = 0;
                        }
                        var getRates = from labor in app2.GetFirstOrDefault().AsEnumerable()
                                       where labor[0].ToString() == CurrentSub
                                       select labor;
                        if (getRates.Count() == 1)
                        {

                            LaborRate = int.Parse(getRates.FirstOrDefault()[iLaborColumnNumber].ToString());
                        }
                        foreach (var columnName in columnNames)
                        {
                            string columnValue = DataItem[ColCount].ToString().Replace(",", "_").Trim();
                            sb.Append(columnValue + ",");
                            if (columnName.ToString().Contains("Equipment") && columnValue != "")
                            {
                                if (getRates.Count() == 1)
                                {

                                    int iEquipColumnNumber = -1;
                                    int equipCurrentColumn = 0;
                                    foreach (var col in rateColumnNames)
                                    {
                                        if (col.ToString() == columnValue.ToString())
                                        {
                                            iEquipColumnNumber = equipCurrentColumn;
                                            break;
                                        }
                                        equipCurrentColumn += 1;

                                    }
                                    if (iEquipColumnNumber >= 0)
                                        EquipmentRate = int.Parse(getRates.FirstOrDefault()[iEquipColumnNumber].ToString());

                                }
                            }
                            ColCount += 1;
                        }
                        sb.Append(LaborRate.ToString() + "," + EquipmentRate.ToString());
                        sb.AppendLine();
                        EquipmentRate = 0;
                    }
                    System.Diagnostics.Trace.WriteLine(sb.ToString());
                    WriteSbToCSV(eventDescr, CurrentSub, sb);




                    Console.WriteLine("going to dispose now...");
                }
            }
            Console.WriteLine("Press return key to end.");
            Console.ReadLine();

        }
        static void WriteSbToCSV(string eventName, string eventSub, StringBuilder sb)
        {
            string sFilename = @"C:\temp\ConsoleTestExcelHelper\SubData\" + eventName + "_" + eventSub + ".csv";
            if (System.IO.File.Exists(sFilename))
                System.IO.File.Delete(sFilename);
            System.IO.File.WriteAllText(@"C:\temp\ConsoleTestExcelHelper\SubData\" + eventName + "_" + eventSub + ".csv", sb.ToString());
        }
        static void TestModel()
        {
            //ExcelHelper.Model.ExcelRows rows = new ExcelHelper.Model.ExcelRows("Row1");
            //rows.Columns.Add(new ExcelHelper.Model.ExcelColumns())
        }
    }
}
