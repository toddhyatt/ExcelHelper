using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ExcelHelper
{
    internal class ExcelWorkSheet : IDisposable
    {
        public string WorkSheetName { get; set; }
        private Worksheet excelWorkSheet { get; set; }

        public ExcelWorkSheet(Workbook excelWorkBook, string ExcelWorkSheetName)
        {
            WorkSheetName = ExcelWorkSheetName;
            if(WorkSheetExists(excelWorkBook,ExcelWorkSheetName))
            {
                excelWorkSheet = (Worksheet)excelWorkBook.Worksheets[ExcelWorkSheetName];
                
            }
            else
            {
                excelWorkSheet = (Worksheet)excelWorkBook.Worksheets.Add();
                excelWorkSheet.Name = WorkSheetName;
            }
            
        }

        public bool WorkSheetExists(Workbook excelWorkBook, string ExcelWorkSheetName)
        {
            bool bReturn = false;
            if (excelWorkBook.Worksheets[ExcelWorkSheetName]!=null)
                bReturn = true;
            return bReturn;
        }


        private bool disposedValue;

        public virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkSheet);

                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ExcelWorkSheet()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
