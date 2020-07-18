using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelHelper
{
    internal class ExcelWorkBook:IDisposable
    {
        private bool disposedValue;

        private Workbook excelWorkBook { get; set; }
        public string FileName { protected set; get; }

        public string TemplateFileName { protected set; get; }

        public List<ExcelWorkSheet> ExcelWorkSheets { get; set; }



        public ExcelWorkBook(Application app,string WorkBookFileName)
        {
            FileName = WorkBookFileName;
            excelWorkBook=FileExists()?app.Workbooks.Open(FileName):app.Workbooks.Add();
            initWorkSheets();
            excelWorkBook.Activate();
        }

        public ExcelWorkBook(Application app, string WorkBookFileName,string TemplateName)
        {
            FileName = WorkBookFileName;
            TemplateFileName = TemplateName;
            excelWorkBook = FileExists() ? app.Workbooks.Open(FileName) : app.Workbooks.Add(TemplateFileName);
            initWorkSheets();
            excelWorkBook.Activate();
        }

        private void initWorkSheets()
        {
            ExcelWorkSheets = new List<ExcelWorkSheet>();
            foreach(Worksheet ws in excelWorkBook.Worksheets)
            {
                ExcelWorkSheets.Add(new ExcelWorkSheet(excelWorkBook, ws.Name));
            }
        }

        public bool FileExists()
        {
            return System.IO.File.Exists(FileName);
        }

        public void Save()
        {
            excelWorkBook.SaveAs(FileName);
        }

        protected void DisposeRanges() { }
        protected void DisposeWorkSheets() 
        {
            foreach(ExcelWorkSheet ws in ExcelWorkSheets)
            {
                ws.Dispose(true);
                
            }
        }

        protected void DisposeWorkBook() 
        {
            excelWorkBook.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkBook);
        }


        public virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    DisposeRanges();
                    DisposeWorkSheets();
                    DisposeWorkBook();

                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ExcelWorkBook()
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
