using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ExcelAppCR.ViewModel
{
    public class PaggingVM : ViewModelBase
    {
        private int _pageSize = 500;
        public int PageSize
        {
            get { return _pageSize; }
            set
            {
                if (_pageSize != value)
                {
                    _pageSize = value;
                    Log.Information("Page Size Changed : " + _pageSize);
                    RaisePropertyChanged(nameof(PageSize));
                }
            }
        }
        private int _rowCount = 0;
        public int RowCount
        {
            get { return _rowCount; }
            set
            {
                if (_rowCount != value)
                {
                    _rowCount = value; 
                    Log.Information("Row Count Changed : " + _rowCount);
                    RaisePropertyChanged(nameof(RowCount));
                }
            }
        } 

        private int _pageIndex = 1; 
        public int PageIndex
        {
            get { return _pageIndex; }
            set
            {
                if (_pageIndex != value)
                {
                    _pageIndex = value;
                    Log.Information("Page Index Changed : " + _pageIndex);
                    RaisePropertyChanged(nameof(PageIndex));
                }
            }
        } 
        public int TotalPages
        {
            get
            {
                if (PageSize == 0) return 0; // Prevent division by zero
                return (int)Math.Ceiling((double)RowCount / PageSize);
            }
        }
    }
}
