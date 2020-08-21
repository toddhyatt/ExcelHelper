using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Model
{
    public class ExcelColumns
    {
        public List<ExcelColumn > Columns { get; set; }

        public string R1C1Name { get; private set; }

        public ExcelColumns(string excelR1C1Name)
        {
            R1C1Name = excelR1C1Name;
            Columns = new List<ExcelColumn>();
        }
        public void AddColumn(string excelR1C1Name, string columnName, string columnValue)
        {
            Columns.Add( new ExcelColumn(excelR1C1Name, columnName, columnValue));
        }
    }
}
