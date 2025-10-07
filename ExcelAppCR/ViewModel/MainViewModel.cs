using ExcelAppCR.Commands;
using ExcelAppCR.Model;
using ExcelAppCR.Service;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.Win32;
using Serilog;
using Microsoft.VisualBasic;
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
        DataLoaded,
        DataSave
    }
    public class MainViewModel : PaggingVM
    {

        private string _filePath;
        public string FilePathEx
        {
            get { return _filePath; }
            set
            {
                _filePath = value;
                Log.Information("File Path :" + _filePath);
                RaisePropertyChanged(nameof(FilePathEx));
            }
        }

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

        private double _lastLoadTime;
        public double LastLoadTime
        {
            get => _lastLoadTime;
            set
            {
                _lastLoadTime = value;
                Log.Information("LoadTime : " + _lastLoadTime);
                RaisePropertyChanged(nameof(LastLoadTime));
            }
        }

        public bool HasData => _dataView != null;

        //Khi sửa ô nào đó ,không ghi ngay vào file mà chỉ ghi log vào list này.
        private List<ExcelFileInfo> _listChange = new List<ExcelFileInfo>();

        // lưu các trang lại sau khi next hoặc previous để không phải load lại từ file
        private Dictionary<int, DataTable> _pageCache = new Dictionary<int, DataTable>();

        private readonly List<string> _addedColumns = new List<string>();

        private ViewState _currentState = ViewState.Empty; // Trạng thái ban đầu
        public ViewState CurrentState
        {
            get => _currentState;
            set
            {
                _currentState = value;
                Log.Information("Current State Changed: {CurrentState}", _currentState);
                RaisePropertyChanged(nameof(CurrentState));
            }
        }


        public MainViewModel()
        {
            InitData();
        }
        ExcelService _excelService;
        public ICommand OpenExcelCommand { get; set; }
        public ICommand SaveFileCommand { get; set; }
        public ICommand NewFile { get; set; }
        public ICommand AddRowCommand { get; set; }
        public ICommand AddColumnCommand { get; set; }
        public void InitData()
        {
            _excelService = new ExcelService();
            OpenExcelCommand = new VfxCommand(OnOpen, () => true);
            SaveFileCommand = new VfxCommand(OnSave, () => true);
            NewFile = new VfxCommand(OnNewFile, () => true);
            AddRowCommand = new VfxCommand(OnAddRow, () => true);
            AddColumnCommand = new VfxCommand(OnAddColumn, () => true);
            CurrentState = ViewState.Empty;
        }

        private void OnAddColumn(object obj)
        {
            if (!HasData)
            {
                MessageBox.Show("Chưa có dữ liệu để thêm cột! Vui lòng tạo file mới", "Warning",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string newColumnName = Interaction.InputBox("Enter new column name:", "Add Column",
                                                            $"NewColumn{_addedColumns.Count + 1}");
                if (string.IsNullOrWhiteSpace(newColumnName))
                    return;

                _addedColumns.Add(newColumnName);

                // cập nhật cache
                foreach (var table in _pageCache.Values)
                    if (!table.Columns.Contains(newColumnName))
                        table.Columns.Add(newColumnName, typeof(string));

                // cập nhật bảng hiện tại
                var currentTable = ExcelData.Table;
                if (!currentTable.Columns.Contains(newColumnName))
                {
                    currentTable.Columns.Add(newColumnName, typeof(string));
                    foreach (DataRow r in currentTable.Rows) r[newColumnName] = "";
                }

                // ép DataGrid refresh schema
                ExcelData = null;
                ExcelData = new DataView(currentTable);

                // schema thay đổi -> dọn state
                _listChange.Clear();
                _pageCache.Clear();

                RefreshPaging();
                Log.Information("Added new column: {ColumnName}", newColumnName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi thêm cột:\n{ex.Message}", "Lỗi",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Add new row to DataTable
        /// </summary>
        private void OnAddRow(object obj)
        {
            if (!HasData)
            {
                MessageBox.Show("Chưa có dữ liệu để thêm dòng! Vui lòng Tạo File mới", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                // lấy DataTable từ DataView hiện tại
                DataTable table = ExcelData.Table;

                var newRow = table.NewRow();
                for (int i = 0; i < ExcelData.Table.Columns.Count; i++)
                {
                    newRow[i] = ""; // Hoặc giá trị mặc định khác
                }
                ExcelData.Table.Rows.Add(newRow);
                TotalRecords++;
                TotalPages = (int)Math.Ceiling((double)TotalRecords / PageSize);

                RaisePropertyChanged(nameof(ExcelData));
                RaisePropertyChanged(nameof(TotalRecords));
                RaisePropertyChanged(nameof(TotalPages));
                (SaveFileCommand as VfxCommand)?.RaiseCanExecuteChanged();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi thêm dòng mới từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnNewFile(object obj)
        {
            CurrentState = ViewState.DataLoaded;

            if (_listChange.Count > 0)
            {
                var result = MessageBox.Show("Dữ liệu hiện tại sẽ bị mất. Bạn có chắc chắn muốn tạo file mới?", "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No)
                    return;
            }
            try
            {
                var table = new DataTable();
                for (int i = 1; i <= 10; i++) table.Columns.Add($"Column {i}", typeof(string));

                for (int i = 0; i < 20; i++)
                {
                    var row = table.NewRow();
                    foreach (DataColumn col in table.Columns) row[col] = "";
                    table.Rows.Add(row);
                }

                ExcelData = table.DefaultView;
                TotalRecords = table.Rows.Count;
                TotalPages = Math.Max(1, (int)Math.Ceiling((double)TotalRecords / PageSize));
                PageIndex = 1;
                _listChange.Clear();
                _pageCache.Clear();
                FilePathEx = string.Empty;
                RefreshPaging();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo file mới từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        protected override async void OnPageSizeChanged(int newSize)
        {
            if (HasData == false)
            {
                _pageCache.Clear();
                _listChange.Clear();
                RefreshPaging();
                return;
            }
            try
            {
                IsProcessing = true;

                // 1) Xóa cache vì page-size thay đổi
                _pageCache.Clear();

                // 2) Nếu đã biết TotalRecords thì tính lại TotalPages
                if (TotalRecords > 0)
                {
                    TotalPages = (int)Math.Ceiling((double)TotalRecords / newSize);
                }

                // 3A) Cách đơn giản: về trang 1
                PageIndex = 1;

                RefreshPaging();
                await LoadPageData();
            }
            finally
            {
                IsProcessing = false;
            }
        }

        /// <summary>
        /// Save file Excel
        /// </summary>
        private async void OnSave(object obj)
        {
            try
            {
                if (!HasData)
                    return;
                if (_listChange.Count == 0)
                {
                    MessageBox.Show("Không có thay đổi nào để lưu!", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (string.IsNullOrWhiteSpace(FilePathEx))
                {
                    await SaveNewFileAsync();
                    return;
                }
                await SaveFileAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi save !", "Errorr", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }


        /// <summary>
        /// Save As New File Excel 
        /// </summary>
        private async Task SaveNewFileAsync()
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Title = "Save Excel File",
                    Filter = "Excel Files|*.xlsx",
                    DefaultExt = ".xlsx",
                };
                if (saveFileDialog.ShowDialog() != true)
                    return;
                IsSaved = true;
                FilePathEx = saveFileDialog.FileName;
                DataTable data = ExcelData.ToTable();
                await _excelService.SaveAsToFile(data, FilePathEx);
                _listChange.Clear();
                _pageCache.Clear();
                MessageBox.Show($"File saved successfully to:\n{FilePathEx}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi save new file ", "Errorr", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                IsSaved = false;
            }
        }

        private async Task SaveFileAsync()
        {
            IsSaved = true;
            try
            {
                await _excelService.SaveToFile(_filePath, _listChange);
                _listChange.Clear();
                _pageCache.Clear();
                _addedColumns.Clear();
                MessageBox.Show($"File saved successfully to:\n{_filePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"File saved successfully to:\n{_filePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            finally
            {
                IsSaved = false;
            }
        }

        /// <summary>
        /// save khi có sự thay đổi trong ô của DataTable
        /// </summary>
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
            _listChange.Clear();
            _pageCache.Clear();
            _addedColumns.Clear();
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Excel File",
                Filter = "Excel Files|*.xlsx",
                DefaultExt = ".xlsx"
            };
            if (openFileDialog.ShowDialog() != true)
                return;
            IsProcessing = true;
            FilePathEx = openFileDialog.FileName;
            try
            {
                CurrentState = ViewState.Loading;
                TotalRecords = (int)await Task.Run(() => _excelService.GetTotalRowCount(FilePathEx));
                Log.Information("Total Record :" + TotalRecords);
                TotalPages = (int)Math.Ceiling((double)TotalRecords / PageSize);
                await LoadPageData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file Excel từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;

            }
        }

        /// <summary>
        /// Load Data from Excel file
        /// </summary>
        public async Task LoadPageData()
        {
            LastLoadTime = 0;
            var stopwatch = Stopwatch.StartNew();
            CurrentState = ViewState.Loading;

            if (_pageCache.ContainsKey(PageIndex))
            {
                ExcelData = _pageCache[PageIndex].DefaultView;
                Log.Information("ExcelData: {ExcelData}", ExcelData);
                RefreshPaging();
                CurrentState = ViewState.DataLoaded;
                return;
            }
            try
            {
                var dataTable = await Task.Run(() => _excelService.LoadExcelPage(FilePathEx, PageIndex, PageSize));
                // lưu vào cache
                _pageCache[PageIndex] = dataTable;
                ExcelData = dataTable.DefaultView;
                RefreshPaging();

            }
            catch (Exception ex)
            {
                CurrentState = ViewState.Empty;
                MessageBox.Show($"Lỗi khi đọc file Excel từ Class MainViewModel:\n {ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                CurrentState = ViewState.DataLoaded;
                stopwatch.Stop();
                LastLoadTime = stopwatch.Elapsed.TotalMilliseconds;
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


