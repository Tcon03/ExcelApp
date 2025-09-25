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
    public class MainViewModel : PaggingVM
    {

        private string _filePath;

        private DataView _dataView;
        public DataView ExcelData
        {
            get { return _dataView; }
            set
            {
                if (_dataView?.Table != null)
                {
                    _dataView.Table.ColumnChanged -= OnCellChanged;
                }
                _dataView = value;

                if (_dataView?.Table != null)
                {
                    _dataView.Table.ColumnChanged += OnCellChanged;
                }

                RaisePropertyChanged(nameof(ExcelData));
                RaisePropertyChanged(nameof(HasData));
            }
        }

        public bool HasData => ExcelData != null && ExcelData.Count > 0;

        // Dùng để lưu trữ thay đổi trên các trang
        private List<ExcelFileInfo> _modifiedCells = new List<ExcelFileInfo>();

        // lưu lại các trang đã được tải để dùng lại khi  chuyển trang

        private Dictionary<int, DataTable> _pageCache = new Dictionary<int, DataTable>();
        public MainViewModel()
        {
            InitData();
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
            if (HasData)
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
                RowCount = 0;
                _filePath = string.Empty;
                RefreshPaging();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo file mới từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async void OnSaveAs(object obj)
        {
            //SaveFileDialog saveFile = new SaveFileDialog
            //{
            //    Title = "Save Excel File",
            //    Filter = "Excel Files|*.xlsx",
            //    DefaultExt = ".xlsx",
            //    FileName = _filePath
            //};
            //Log.Information("File Path : " + saveFile.FileName);
            //if (saveFile.ShowDialog() != true)
            //    return;
            //string path = saveFile.FileName;
        }
        private async void OnSave(object obj)
        {



            try
            {
                await _excelService.SaveToFile(_filePath, _modifiedCells);
                _modifiedCells.Clear();
                _pageCache.Clear();
                MessageBox.Show($"File saved successfully to:\n{_filePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                await LoadPageData();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi save ! ", "Errorr", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }



        private void OnCellChanged(object sender, DataColumnChangeEventArgs e)
        {
            var rowIndexOnPage = e.Row.Table.Rows.IndexOf(e.Row);
            Log.Information("Row Index On Page: {RowIndexOnPage}", rowIndexOnPage);
            var absoluteRowIndex = ((PageIndex - 1) * PageSize) + rowIndexOnPage + 2; 
            Log.Information("Absolute Row Index: {AbsoluteRowIndex}", absoluteRowIndex);
            var columnIndex = e.Column.Ordinal + 1; 
            Log.Information("Column Index: {ColumnIndex}", columnIndex);

               var change = new ExcelFileInfo
            {
                RowIndex = absoluteRowIndex,
                ColumnIndex = columnIndex,
                NewValue = e.ProposedValue
            };

            _modifiedCells.RemoveAll(c => c.RowIndex == change.RowIndex && c.ColumnIndex == change.ColumnIndex);
            _modifiedCells.Add(change);

            Log.Information("Cell changed: Row {Row}, Col {Col}", change.RowIndex, change.ColumnIndex);
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
            IsProcessing = true;

            try
            {
                RowCount = (int)await Task.Run(() => _excelService.GetTotalRowCount(_filePath));
                Log.Information("RowCount :" + RowCount);
                TotalPages = (int)Math.Ceiling((double)RowCount / PageSize);
                await LoadPageData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file Excel từ Class MainViewModel:\n{ex.Message}",
                                 "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                RefreshPaging();
                IsProcessing = false;
            }
        }

        /// <summary>
        /// Load dữ liệu trang hiện tại từ file Excel
        /// </summary>
        /// <returns></returns>
        public async Task LoadPageData()
        {
            if (string.IsNullOrEmpty(_filePath))
                return;


            try
            {
                // Get data for 
                var dataTable = await Task.Run(() => _excelService.LoadExcelPage(_filePath, PageIndex, PageSize));
                _pageCache[PageIndex] = dataTable;
                ExcelData = dataTable.DefaultView;
                RefreshPaging();
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

        public override async void OnNextPage(object obj)
        {
            if (CanGoNextPage())
            {
                PageIndex++;
                await LoadPageData();
            }
        }
        public override async void OnPreviousPage(object obj)
        {
            if (CanGoPreviousPage())
            {
                PageIndex--;
                await LoadPageData();
            }
        }
    }
}


