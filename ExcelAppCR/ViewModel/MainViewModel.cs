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
using Microsoft.Win32;
using ExcelAppCR.Commands; 
using Serilog;


namespace ExcelAppCR.ViewModel
{
    public class MainViewModel : ViewModelBase
    {
        private DataView _dataView;
        public DataView ExcelData
        {
            get { return _dataView; }
            set
            {
                _dataView = value;
                Log.Information("value DataView :" + _dataView);
                RaisePropertyChange(nameof(ExcelData));
            }
        }
        //ExcelService excelService;
        public MainViewModel()
        {
            InitData();
        }

        public ICommand OpenExcelCommand { get; set; }
        public ICommand SaveFileCommand { get; set; }
        public void InitData()
        {
            //excelService = new ExcelService();
            OpenExcelCommand = new VfxCommand(OnOpen, () => true);
            SaveFileCommand = new VfxCommand(OnSave, () => true);
        }



        private void OnSave(object obj)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.DefaultExt = ".xlsx";
            saveFile.Title = "Save Excel File";

        }
        private void OnOpen(object obj)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                if (openFileDialog.ShowDialog() != true)
                    return;
                //var fs = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                //Log.Information(" ----- Open File -----  : {FileName}", openFileDialog.FileName);
                //var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(fs);
                //var conf = new ExcelDataSetConfiguration
                //{
                //    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                //    { 
                //        UseHeaderRow = true
                //    }
                //};
                //var ds = reader.AsDataSet(conf);
                //if (ds.Tables.Count == 0)
                //{
                //    ExcelData = null;
                //    return;
                //}
                //DataTable sheet = ds.Tables[0];

                //var data = excelService.ReadExcelData(openFileDialog.FileName);
                //ExcelData = data.DefaultView;
                //MessageBox.Show("Load Excel File Success" + data.Rows.Count, "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                //Log.Information(" ----- Load Excel File Success -----  : {Row}", data.Rows.Count);

            }
            catch (Exception ex)
            {
                Log.Error(" ----- Error Open File -----  : {Error}", ex.Message);
            }
        }
    }
}

