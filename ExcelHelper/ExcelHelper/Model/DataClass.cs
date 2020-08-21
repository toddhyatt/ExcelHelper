using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace ExcelHelper.Model
{
    public class DataClass
    {
        public DataTable dt { get; set; }
        public DataClass()
        {
            dt = new DataTable();
        }
    }
}
