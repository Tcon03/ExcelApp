using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAppCR.Model
{
    public class ExcelFileInfo
    {
        public DataTable FullData { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public string SheetName { get; set; }
    }
}
