using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Model
{
    public class ExcelRows
    {
        public ExcelColumns Columns { get; set; }

        public string R1C1Name { get; private set; }

        public ExcelRows(string excelR1C1Name)
        {
            Columns = new ExcelColumns(excelR1C1Name);
        }
    }
}
