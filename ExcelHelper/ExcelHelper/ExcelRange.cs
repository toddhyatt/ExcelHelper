using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Net.Http.Headers;

namespace ExcelHelper
{
    internal class ExcelRange : IDisposable
    {

        public string RangeName { get; set; }
        
        private Range excelRange { get; set; }

        internal System.Data.DataTable rangeDT { get; set; }

        
        private bool disposedValue;

        public ExcelRange(Worksheet excelWorkSheet,string excelRangeName)
        {
            RangeName = excelRangeName;
            excelRange = excelWorkSheet.Range[RangeName];
            //excelRange=excelWorkSheet.
            rangeDT=BuildDataTable();
        }

        internal System.Data.DataTable BuildDataTable()
        {
            string CurrentLocation = string.Empty;
            int rowcount = 0;
            int colCount = 0;
            DataRow dr;
            System.Data.DataTable dt = new System.Data.DataTable(RangeName);
            foreach (Range excelRow in excelRange.Rows)
            {
                if (rowcount == 0)
                {
                    //build columns
                    foreach (Range excelColumn in excelRange.Columns)
                    {
                        dt.Columns.Add(excelColumn.Address);
                    }
                }
                dr = dt.NewRow();
                foreach (Range excelColumn in excelRow.Columns)
                {
                    //set current location
                    if(rowcount >= 8 && colCount == 1 && excelColumn.Value != null)
                        CurrentLocation = excelColumn.Value.ToString();
                    //Replace column name with location
                    if (rowcount == 5 && colCount == 1 && excelRange.Worksheet.Application.ActiveWorkbook.Name.Contains("Event"))
                        dr[colCount] = "Location";
                    else if (rowcount > 8 && colCount == 1)
                        dr[colCount] = CurrentLocation;
                    else
                        dr[colCount] = excelColumn.Value;
                    colCount = colCount + 1;
                }
                dt.Rows.Add(dr);
                rowcount = dt.Rows.Count;
                colCount = 0;
            }
            return dt;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~ExcelRange()
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
