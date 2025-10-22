using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp.Model
{
    public class ExcelFileInfo
    {
        public int RowIndex { get; set; }    // Số thứ tự dòng (trong Excel)
        public int ColumnIndex { get; set; } // Số thứ tự cột (trong Excel)
        public object NewValue { get; set; }   // Giá trị mới
    }
}
