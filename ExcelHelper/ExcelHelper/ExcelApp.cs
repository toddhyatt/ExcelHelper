using System;
using System.Dynamic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper
{
    public class ExcelApp : IDisposable
    {
        private bool disposedValue;

        private Application excelApp { get; set; }
        private List<ExcelWorkBook> excelWorkBooksList{ get; set; }
        
        public ExcelApp()
        {
            InitClass();
        }

        public void OpenWorkBook(string FileName)
        {
            excelWorkBooksList.Add(new ExcelWorkBook(excelApp, FileName));
        }

        public void OpenWorkBookTemplate(string FileName,string TemplateFileName)
        {
            excelWorkBooksList.Add(new ExcelWorkBook(excelApp, FileName,TemplateFileName));
        }


        public void SaveWorkBooks()
        {
            var q = from wBooks in excelWorkBooksList
                    select wBooks;
            foreach(ExcelWorkBook wBook in q)
            {
                SaveWorkBook(wBook.FileName);
            }

        }
        public void SaveWorkBook(string FileName)
        {
            var q = from wBooks in excelWorkBooksList
                    where wBooks.FileName == FileName
                    select wBooks;
            q.FirstOrDefault<ExcelWorkBook>().Save();
        }

        protected void InitClass()
        {
            excelApp = new Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            excelWorkBooksList = new List<ExcelWorkBook>();
        }
        

        protected void DisposeWorkBooks() {
            foreach(ExcelWorkBook excelWorkBook in excelWorkBooksList)
            {
                excelWorkBook.Dispose(true);
            }
            
        }
        protected void DisposeApplication() 
        {
            if(excelApp!=null)
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                //excelApp = null;
            }
        }


        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                DisposeWorkBooks();
                DisposeApplication();

                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ExcelHelper()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        void IDisposable.Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            //GC.SuppressFinalize(this);
        }
        ~ExcelApp()
        {
            this.Dispose(true);
        }

    }
}
