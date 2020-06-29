using System;
using System.Dynamic;
using System.IO;

namespace ExcelHelper
{
    public class ExcelApp : IDisposable
    {
        private bool disposedValue;
        public string FileName { protected set; get; }

        public ExcelApp(string sFileName )
        {
            FileName = sFileName;
        }

        public bool DoesFileExist()
        {
            return System.IO.File.Exists(FileName);
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
        // ~ExcelHelper()
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
