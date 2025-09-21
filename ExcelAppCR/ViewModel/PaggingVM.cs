using ExcelAppCR.Commands;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelAppCR.ViewModel
{
    public abstract class PaggingVM : ViewModelBase
    {
        public PaggingVM()
        {
            NextPageCommand = new VfxCommand(OnNextPage, CanGoNextPage);
            PreviousPageCommand = new VfxCommand(OnPreviousPage, CanGoPreviousPage);
        }

        public bool CanGoPreviousPage()
        {
            if (PageIndex > 1)
                return true;
            return false;
        }

        public void OnPreviousPage(object obj)
        {
            if (PageIndex > 1)
                PageIndex--;
            Log.Information("Navigated to Previous Page: {PageIndex}", PageIndex);
        }

        public bool CanGoNextPage()
        {
            if (PageIndex < TotalPages)
                return true;
            return false;
        }

        public void OnNextPage(object obj)
        {
            if (PageIndex < TotalPages)
                PageIndex++;
            Log.Information("Navigated to Next Page: {PageIndex}", PageIndex);
        }

        private int _pageSize = 1000;
        public int PageSize
        {
            get { return _pageSize; }
            set
            {
                if (_pageSize != value)
                {
                    _pageSize = value;
                    Log.Information("Page Size Changed : {PageSize}", _pageSize);
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
                    Log.Information("Row Count Changed : {RowCount}", _rowCount);
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
                    Log.Information("Page Index Changed : {PageIndex}", _pageIndex);
                    RaisePropertyChanged(nameof(PageIndex));
                    LoadPageData();
                }
            }
        }
        private int _totalPages = 0;
        public int TotalPages
        {
            get
            {
                return _totalPages;
            }
            set
            {
                if (_totalPages != value)
                {
                    _totalPages = value;
                    Log.Information("Total Pages Changed : {TotalPages}", _totalPages);
                    RaisePropertyChanged(nameof(TotalPages));
                }
            }
        }

        private bool _isProcessing;
        public bool IsProcessing
        {
            get => _isProcessing;
            set
            {
                _isProcessing = value;
                Log.Information("IsProcessing Changed : {IsProcessing}", _isProcessing);
                RaisePropertyChanged(nameof(IsNotProcessing));
            }
        }
        public void RefreshPaging()
        {
            (NextPageCommand as VfxCommand)?.RaiseCanExecuteChanged();
            (PreviousPageCommand as VfxCommand)?.RaiseCanExecuteChanged();
        }
        public bool IsNotProcessing => !_isProcessing;

        public ICommand NextPageCommand { get; set; }
        public ICommand PreviousPageCommand { get; set; }
        protected abstract void LoadPageData();
    }
}
