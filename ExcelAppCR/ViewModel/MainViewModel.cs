using ExcelAppCR.Commands;
using ExcelAppCR.Model;
using ExcelAppCR.Service;
using Microsoft.Win32;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;


namespace ExcelAppCR.ViewModel
{
    public enum ViewState
    {
        Empty,
        Loading,
        DataLoaded
    }
    public class MainViewModel : PaggingVM
    {

        private string _filePath;

        private DataView _dataView;
        public DataView ExcelData
        {
            get { return _dataView; }
            set
            {
                // hủy đăng ký sự kiện cũ 
                if (_dataView?.Table != null)
                {
                    _dataView.Table.ColumnChanged -= OnCellChanged;
                }
                // gán giá trị mới
                _dataView = value;

                // đăng ký sự kiện mới
                if (_dataView?.Table != null)
                {
                    _dataView.Table.ColumnChanged += OnCellChanged;
                }

                RaisePropertyChanged(nameof(ExcelData));
                RaisePropertyChanged(nameof(HasData));
            }
        }

        public bool HasData => TotalRecords > 0;
        public bool HasDataTable => _dataView != null;

        //Khi sửa ô nào đó ,không ghi ngay vào file mà chỉ ghi log vào list này.
        private List<ExcelFileInfo> _listChange = new List<ExcelFileInfo>();

        // lưu các trang lại sau khi next hoặc previous để không phải load lại từ file
        private Dictionary<int, DataTable> _pageCache = new Dictionary<int, DataTable>();
        public MainViewModel()
        {
            InitData();
        }



        private ViewState _currentState = ViewState.Empty; // Trạng thái ban đầu
        public ViewState CurrentState
        {
            get => _currentState;
            set
            {
                _currentState = value;
                RaisePropertyChanged(nameof(CurrentState));
            }
        }



        ExcelService _excelService;
        public ICommand OpenExcelCommand { get; set; }
        public ICommand SaveFileCommand { get; set; }
        public ICommand NewFile { get; set; }
        public void InitData()
        {
            _excelService = new ExcelService();
            OpenExcelCommand = new VfxCommand(OnOpen, () => true);
            SaveFileCommand = new VfxCommand(OnSave, () => true);
            NewFile = new VfxCommand(OnNewFile, () => true);
        }

        private void OnNewFile(object obj)
        {
            if (HasDataTable)
            {
                var result = MessageBox.Show("Dữ liệu hiện tại sẽ bị mất. Bạn có chắc chắn muốn tạo file mới?", "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                    return;
            }
            try
            {
                var table = new DataTable();
                table.Columns.Add("Col1");
                table.Columns.Add("Col2");
                table.Columns.Add("Col3");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                table.Rows.Add("", "", "");
                ExcelData = table.DefaultView;
                PageIndex = 1;
                TotalPages = 0;
                TotalRecords = 0;
                _filePath = string.Empty;
                RefreshPaging();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo file mới từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        private async void OnSave(object obj)
        {
            try
            {
                await _excelService.SaveToFile(_filePath, _listChange);
                _listChange.Clear();
                _pageCache.Clear();
                MessageBox.Show($"File saved successfully to:\n{_filePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);


            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi save ! ", "Errorr", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }



        private void OnCellChanged(object sender, DataColumnChangeEventArgs e)
        {
            //1 Tính Toán vị trí tuyệt đối của ô trong file Excel
            var rowIndexOnPage = e.Row.Table.Rows.IndexOf(e.Row);
            Log.Information("Row Index On Page: {RowIndexOnPage}", rowIndexOnPage);

            var absoluteRowIndex = ((PageIndex - 1) * PageSize) + rowIndexOnPage + 2;
            Log.Information("Absolute Row Index: {AbsoluteRowIndex}", absoluteRowIndex);
            var columnIndex = e.Column.Ordinal + 1;
            Log.Information("Column Index: {ColumnIndex}", columnIndex);

            // 2. tạo đối tượng ExcelFileInfo để lưu thông tin thay đổi
            var _change = new ExcelFileInfo
            {
                RowIndex = absoluteRowIndex,
                ColumnIndex = columnIndex,
                NewValue = e.ProposedValue
            };

            //3. kiểm tra và cập nhật thay đổi mới vào danh sách 
            _listChange.RemoveAll(c => c.RowIndex == _change.RowIndex && c.ColumnIndex == _change.ColumnIndex);
            _listChange.Add(_change);

            Log.Information("Cell changed: Row {Row}, Col {Col}", _change.RowIndex, _change.ColumnIndex);

            (SaveFileCommand as VfxCommand)?.RaiseCanExecuteChanged();
        }
        /// <summary>
        /// Open file Excel and load data with paging
        /// </summary>
        private async void OnOpen(object obj)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Excel File",
                Filter = "Excel Files|*.xlsx",
                DefaultExt = ".xlsx"
            };
            if (openFileDialog.ShowDialog() != true)
                return;
            _filePath = openFileDialog.FileName;
            CurrentState = ViewState.Empty;
            PageIndex = 1;
            try
            {
                TotalRecords = (int)await Task.Run(() => _excelService.GetTotalRowCount(_filePath));
                Log.Information("Total Record :" + TotalRecords);
                TotalPages = (int)Math.Ceiling((double)TotalRecords / PageSize);
                await LoadPageData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file Excel từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        /// <summary>
        /// Load Data from Excel file
        /// </summary>
        public async Task LoadPageData()
        {
            CurrentState = ViewState.Loading;
            if (_pageCache.ContainsKey(PageIndex))
            {
                ExcelData = _pageCache[PageIndex].DefaultView;
                Log.Information("ExcelDât: {ExcelData}", ExcelData);
                RefreshPaging();
                return;
            }
            try
            {
                // Get data for 
                var dataTable = await Task.Run(() => _excelService.LoadExcelPage(_filePath, PageIndex, PageSize));
                // lưu vào cache
                _pageCache[PageIndex] = dataTable;
                ExcelData = dataTable.DefaultView;
                RefreshPaging();
                CurrentState = ViewState.DataLoaded;
            }
            catch (Exception ex)
            {
                CurrentState = ViewState.Empty;
                MessageBox.Show($"Lỗi khi đọc file Excel từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }


        public override async void OnNextPage(object obj)
        {
            try
            {
                IsProcessing = true;
                if (CanGoNextPage())
                {
                    PageIndex++;
                    await LoadPageData();
                }

            }
            finally
            {
                IsProcessing = false;
            }
        }


        public override async void OnPreviousPage(object obj)
        {
            try
            {
                IsProcessing = true;
                if (CanGoPreviousPage())
                {
                    PageIndex--;
                    await LoadPageData();
                }
            }
            finally
            {
                IsProcessing = false;
            }
        }
    }
}


