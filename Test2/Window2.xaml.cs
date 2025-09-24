using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
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
using System.Windows.Shapes;

namespace Test2
{
    /// <summary>
    /// Interaction logic for Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {

        private DataTable _excelData;

        public Window2()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("Nguyen Truong");
        }

        // Nút Open - Mở file Excel
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Title = "Chọn file Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls",
                DefaultExt = ".xlsx"
            };

            if (openDialog.ShowDialog() != true)
                return;

            try
            {
                // Đọc file Excel
                _excelData = ReadExcelFile(openDialog.FileName);

                // Hiển thị dữ liệu trên DataGrid
                DataGridExcel.ItemsSource = _excelData.DefaultView;

                // Cập nhật tên file
                TxtFileName.Text = $"File: {System.IO.Path.GetFileName(openDialog.FileName)}";

                MessageBox.Show("Mở file thành công!", "Thông báo",
                               MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi mở file:\n{ex.Message}", "Lỗi",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DataTable ReadExcelFile(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Lấy worksheet đầu tiên
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Kiểm tra worksheet có dữ liệu không
                if (worksheet.Dimension == null)
                {
                    throw new Exception("File Excel không có dữ liệu!");
                }

                // Tạo cột cho DataTable (từ hàng đầu tiên)
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                    dataTable.Columns.Add(columnName);
                }

                // Đọc dữ liệu từ hàng thứ 2 trở đi
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        // Nút Save - Lưu file Excel
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (_excelData == null || _excelData.Rows.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu để lưu!", "Thông báo",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Title = "Lưu file Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                FileName = "DuLieuDaSua.xlsx"
            };

            if (saveDialog.ShowDialog() != true)
                return;

            try
            {
                // Lưu file Excel
                SaveExcelFile(_excelData, saveDialog.FileName);

                MessageBox.Show($"Lưu file thành công!\nĐường dẫn: {saveDialog.FileName}",
                               "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu file:\n{ex.Message}", "Lỗi",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveExcelFile(DataTable dataTable, string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Tạo worksheet mới
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");

                // Ghi header (tên cột)
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                }

                // Ghi dữ liệu
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                // Tự động điều chỉnh độ rộng cột
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Lưu file
                package.Save();
            }
        }
    }
}
