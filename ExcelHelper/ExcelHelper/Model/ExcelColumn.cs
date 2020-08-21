using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelHelper.Model
{
    public class ExcelColumn
    {
        public string R1C1Name { get; private set; }
        public string ColumnName { get; private set; }
        public bool HasChange { get; private set; }

        public string Value {

            get { return Value; }
            set {
                if (Value!=null && Value != value)
                    HasChange = true;
                Value = value;
            }

        }
        public ExcelColumn(string excelR1C1Name,string columnName,string strValue)
        {
            R1C1Name = excelR1C1Name;
            ColumnName = columnName;
            Value = strValue;

        }
    }
}
