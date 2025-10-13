using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp.Model
{
    public class CellChange
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public object NewValue { get; set; }
    }
}
