using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp.ViewModel
{
    public class PaggingVM : ViewModelBase
    {
        private int _pageSize = 10;
        public int PageSize
        {
            get => _pageSize;
            set
            {
                if (value > 0)
                {
                    _pageSize = value;
                    RaisePropertyChanged(nameof(PageSize));
                }
            }
        }

        private int _totalRecords = 0;
        public int TotalRecords
        {
            get => _totalRecords;
            set
            {
                if (value >= 0)
                {
                    _totalRecords = value;
                    
                    RaisePropertyChanged(nameof(TotalRecords));
                    RaisePropertyChanged(nameof(TotalPages));
                }
            }
        }
        public int TotalPages => (int)Math.Ceiling((double)TotalRecords / PageSize);


        private int _pageIndex = 1;
        public int PageIndex
        {
            get => _pageIndex;
            set
            {
                if (value > 0 && value <= TotalRecords)
                {
                    _pageIndex = value;
                    RaisePropertyChanged(nameof(PageIndex));
                }
            }
        }


        private double _lastLoadTime;
        public double LastLoadTime
        {
            get => _lastLoadTime;
            set
            {
                _lastLoadTime = value;
                RaisePropertyChanged(nameof(LastLoadTime));
            }
        }
    }
}
