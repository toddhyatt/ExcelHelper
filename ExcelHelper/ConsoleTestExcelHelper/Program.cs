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
            ProcessAllFiles();
            Console.WriteLine("Press return key to end.");
            Console.ReadLine();

        }

        static void ProcessAllFiles()
        {
            string reportsFolder = @"C:\Users\toddh\OneDrive\villagegreenlandscapes\streetsmartreports";
            foreach (string file in System.IO.Directory.GetFiles(reportsFolder,"*17*"))
            {
                ProcessFile(file);
            }
        }

        static void ProcessFile(string filename)
        {
            using (ExcelHelper.ExcelApp app = new ExcelHelper.ExcelApp())
            {
                using (ExcelHelper.ExcelApp app2 = new ExcelHelper.ExcelApp())
                {
                    string templateName = @"C:\temp\ConsoleTestExcelHelper\Template.xltx";
                    //string templateName = "Template.xltx";
                    //app.OpenWorkBook(@"C:\temp\Event #5 - FULL EVENT REPORT.xls");
                    app.OpenWorkBook(filename);
                    System.IO.FileInfo fi = new System.IO.FileInfo(filename);

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
                    StringBuilder sbLog = new StringBuilder();
                    var eventDescr = app.GetFirstOrDefault().AsEnumerable().FirstOrDefault()[0].ToString();
                    var runBy = app.GetFirstOrDefault().AsEnumerable().FirstOrDefault()[1].ToString();
                    System.Diagnostics.Trace.WriteLine(eventDescr.ToString());
                    var columnNames = from myRow in app.GetFirstOrDefault().Rows[5].ItemArray
                                      select myRow;
                    string commaDelimitedColumnNames = "";
                    foreach (var columnName in columnNames)
                    {
                        commaDelimitedColumnNames += columnName.ToString() + ",";
                        System.Diagnostics.Trace.WriteLine(columnName.ToString());
                    }
                    commaDelimitedColumnNames += "LaborRate,EquipmentRate,SaltBagRate,SaltTonRate,LaborCalc,LaborNotes,ProductCalc,ProductNotes";
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
                    int SaltBagRate = 0;
                    int SaltTonRate = 0;

                    StringBuilder sb = new StringBuilder();

                    var rateColumnNames = from myRow in app2.GetFirstOrDefault().Rows[0].ItemArray
                                          select myRow;
                    
                    int iLaborColumnNumber = -1;
                    int iSaltBagsColumn = -1;
                    int iSaltTonsColumn = -1;

                    int currentColumn = 0;
                    foreach (var col in rateColumnNames)
                    {
                        if (col.ToString() == "Labor")
                        {
                            iLaborColumnNumber = currentColumn;
                            
                        }
                        if (col.ToString() == "Salt Bags")
                        {
                            iSaltBagsColumn = currentColumn;
                        }
                        if (col.ToString() == "Salt Tons")
                        {
                            iSaltTonsColumn = currentColumn;
                        }
                        if (iLaborColumnNumber > 0 && iSaltBagsColumn > 0 && iSaltTonsColumn > 0)
                            break;
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
                            SaltBagRate = 0;
                            SaltTonRate = 0;
                        }

                        var getRates = from labor in app2.GetFirstOrDefault().AsEnumerable()
                                       where labor[0].ToString() == CurrentSub
                                       select labor;
                        if (getRates.Count() == 1)
                        {

                            LaborRate = int.Parse(getRates.FirstOrDefault()[iLaborColumnNumber].ToString());
                            SaltBagRate = int.Parse(getRates.FirstOrDefault()[iSaltBagsColumn].ToString());
                            SaltTonRate = int.Parse(getRates.FirstOrDefault()[iSaltTonsColumn].ToString());
                        }
                        string JobType = "";
                        float Duration = 0;
                        float EmployeeCount = 0;
                        float ProductQty = 0;
                        string JobId = "";
                        foreach (var columnName in columnNames)
                        {
                            string searchFor = ",";
                            string replaceWith = "_";
                            string[] equips;
                            string[] equipRates;
                            string columnValue;
                            if (columnName.ToString().Contains("Equipment"))
                            {
                                columnValue= DataItem[ColCount].ToString().Trim();
                            }
                            else
                            { 
                               columnValue = DataItem[ColCount].ToString().Replace(searchFor, replaceWith).Trim();
                            }
                            equips = columnValue.Split(',');
                            equipRates = new string[equips.Length];
                            if (equips.Length > 1)
                                System.Diagnostics.Trace.WriteLine(columnValue);
                            if (columnName.ToString() == "Job ID")
                                JobId = columnValue;
                            if (columnName.ToString() == "Job Type Name")
                                JobType = columnValue;
                            if (columnName.ToString().Contains("Employees") && columnValue != "")
                            {
                                if (!float.TryParse(columnValue, out EmployeeCount))
                                { 
                                    EmployeeCount = 0;
                                    LogIssue(sbLog,fi.Name,CurrentSub, JobId, "Invalid EmployeeCount", columnValue);
                                }
                            }
                            if (columnName.ToString() == "Duration")
                            {
                                if (!float.TryParse(columnValue, out Duration))
                                {
                                    Duration = 0;
                                    LogIssue(sbLog,fi.Name, CurrentSub, JobId, "Invalid Duration", columnValue);
                                }
                              
                            }         
                            if (columnName.ToString().Contains("SaltBags") && columnValue != "")
                            {
                                if (!float.TryParse(columnValue, out ProductQty))
                                {
                                    ProductQty = 0;
                                    LogIssue(sbLog,fi.Name, CurrentSub, JobId, "Invalid SaltBags", columnValue);
                                }
                            }
                            
                            if (columnName.ToString().Contains("SaltApplied") && columnValue != "")
                            {
                                if (!float.TryParse(columnValue, out ProductQty))
                                { 
                                    ProductQty = 0;
                                    LogIssue(sbLog,fi.Name, CurrentSub, JobId, "Invalid SaltApplied", columnValue);
                                }
                            }
                            
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
                                    {   
                                        if (columnName.ToString().Contains("Haul"))
                                            System.Diagnostics.Trace.WriteLine(CurrentSub + ": Sub has Haul Data");
                                        EquipmentRate = int.Parse(getRates.FirstOrDefault()[iEquipColumnNumber].ToString());
                                    }

                                }
                            }
                            ColCount += 1;
                        }
                        string LaborCalcNotes = "";
                        string ProductCalcNotes = "";
                        float LaborCalc = CalcLabor(JobType, Duration, EmployeeCount, LaborRate, EquipmentRate, ref LaborCalcNotes);
                        if (LaborCalcNotes != "")
                            LogIssue(sbLog, fi.Name, CurrentSub, JobId, "Issues With Labor Calc", LaborCalcNotes);
                        float ProductCalc = CalcProduct(JobType, ProductQty, SaltBagRate, SaltTonRate, ref ProductCalcNotes);
                        if (ProductCalcNotes != "")
                            LogIssue(sbLog, fi.Name, CurrentSub, JobId, "Issues With Product Calc", ProductCalcNotes);
                        sb.Append(LaborRate.ToString() + "," + EquipmentRate.ToString() + "," + SaltBagRate + "," + SaltTonRate + "," + LaborCalc.ToString() + "," + LaborCalcNotes + "," + ProductCalc.ToString() + "," + ProductCalcNotes);
                        sb.AppendLine();
                        EquipmentRate = 0;
                    }
                    System.Diagnostics.Trace.WriteLine(sb.ToString());
                    WriteSbToCSV(eventDescr, CurrentSub, sb);
                    if(sbLog.ToString()!="")
                        WriteSbToCSV(eventDescr, "IssueLog", sbLog);




                    Console.WriteLine("going to dispose now...");
                }
            }

        }
        static void LogIssue(StringBuilder logBuilder,string File,string SubName ,string Job,string observation,string colValue)
        {
            logBuilder.AppendLine(string.Format("Filename:{0} SubName: {1} JobId:{2} Message:{3} ColumnValue:{4}", File, SubName, Job, observation,colValue));
        }
        static float CalcProduct(string JobType, float ProductQty, float SaltBagRate, float SaltTonRate, ref string ProductCalcNotes)
        {
            StringBuilder productNotes = new StringBuilder();
            float product = 0;

            if (JobType == "WINTER - Salting")
            {
                if (SaltTonRate == 0)
                    productNotes.Append(" NOTONRATE Entered Calc is 0");
                if (ProductQty == 0)
                    productNotes.Append(" INVALIDQTY Entered Calc is 0");

                product = ProductQty*SaltTonRate;
            }
            if (JobType == "WINTER - Walks" && ProductQty > 0)
            {
                if (SaltBagRate == 0)
                    productNotes.Append(" NOBAGRATE Entered Calc is 0");
                product = ProductQty * SaltBagRate;
            }

            return product;
        }
        static float CalcLabor(string Job,float Duration,float EmployeeCount,float LaborRate, float EquipRate,ref string Notes)
        {
            StringBuilder laborNotes = new StringBuilder();
            float labor = -1;
            if (Duration == 0)
                laborNotes.Append(" NOTIME Entered Calc is 0");
            if (EmployeeCount == 0)
                laborNotes.Append(" NOEMPLOYEE Entered Calc is 0");
            if (Job == "WINTER - Hauling" || Job == "WINTER - Plowing" || Job=="WINTER - Plowing Cleanup" )
            {
                if(EquipRate==0)
                    laborNotes.Append(" NOEQUIPRATE Entered Calc is 0");
                if (EmployeeCount > 1)
                {
                    laborNotes.Append(" TOMANYEMPLOYEE Entered Set To 1");
                    EmployeeCount = 1;
                }
                float laborPerMinute = EquipRate / 60;
                labor = Duration * laborPerMinute * EmployeeCount;
            }
            if (Job=="WINTER - Walks" || Job == "WINTER - Salting")
            {
                if (LaborRate == 0)
                    laborNotes.Append(" NOLABORRATE Entered Calc is 0");
                float laborPerMinute = LaborRate / 60;
                labor = Duration * laborPerMinute * EmployeeCount;
            }
            Notes = laborNotes.ToString();
            if(labor==-1)
            {
                labor = 0;
                laborNotes.Append(" NOLABORCALC For Job Type -" + Job);
            }
            return labor;
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
