using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace test3Ktem
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string _filePath;
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Import_click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx";
            if (openFileDialog.ShowDialog() != true)
                return;
            _filePath = openFileDialog.FileName;
            var fileInfo = new System.IO.FileInfo(_filePath);
            using (var packege = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet exWorksheet = packege.Workbook.Worksheets[0]; 

                int readStartRow = exWorksheet.Dimension.Start.Row+1;
                Debug.WriteLine("startRow :" + readStartRow);

                int readEndRow = exWorksheet.Dimension.End.Row;
                Debug.WriteLine("endRow :" + readEndRow); 

                for (int i = readStartRow; i <= readEndRow; i++)
                {
                    string rowData = "";
                    for (int j = exWorksheet.Dimension.Start.Column; j <= exWorksheet.Dimension.End.Column; j++)
                    {
                        rowData += exWorksheet.Cells[i, j].Text + "\t";
                    }
                    Debug.WriteLine("rowData :" + rowData);
                }
            }
        }
    }
}
