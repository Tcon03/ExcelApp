using ExcelAppCR.Commands;
using ExcelAppCR.Service;
using Microsoft.Win32;
using Serilog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;


namespace ExcelAppCR.ViewModel
{
    public class MainViewModel : PaggingVM
    {

        string _filePath;

        private DataView _dataView;
        public DataView ExcelData
        {
            get { return _dataView; }
            set
            {
                _dataView = value;
                Log.Information("ExcelData set. Rows={Rows}, Columns={Cols}");

                RaisePropertyChanged(nameof(ExcelData));
            }
        }


        ExcelService _excelService;
        public ICommand OpenExcelCommand { get; set; }
        public ICommand SaveFileCommand { get; set; }
        public void InitData()
        {
            _excelService = new ExcelService();
            OpenExcelCommand = new VfxCommand(OnOpen, () => true);
            SaveFileCommand = new VfxCommand(OnSave, () => true);
        }
        public MainViewModel()
        {
            InitData();
        }


        private void OnSave(object obj)
        {
            SaveFileDialog saveFile = new SaveFileDialog
            {
                Title = "Save Excel File",
                Filter = "Excel Files|*.xlsx",
                DefaultExt = ".xlsx"
            };

        }
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

        protected async override Task LoadPageData()
        {
            TotalPages = (int)Math.Ceiling((double)RowCount / PageSize);
            Log.Information("Total Pages: {TotalPages}", TotalPages);

            var dataTable = await Task.Run(() => _excelService.LoadExcelPage(_filePath, PageIndex, PageSize));
            ExcelData = dataTable.DefaultView;

        }

    }
}


