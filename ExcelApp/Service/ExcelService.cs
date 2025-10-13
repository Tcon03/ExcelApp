using ExcelApp.Model;
using OfficeOpenXml;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelApp.Service
{
    public class ExcelService
    {

        public ExcelService()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Nguyen Truong");
        }

        /// <summary>
        /// Load file Excel theo trang (pagination)
        /// </summary>
        public DataTable LoadExcelPage(string filePath, int pageIndex, int pageSize)
        {
            try
            {
                //1. Khởi tạo
                var dataTable = new DataTable();

                var fileInfo = new FileInfo(filePath);
                Log.Information("Getting total row count from file: {FilePath}", filePath);

                // 2 . mở file excel và đọc dữ liệu
                using (var package = new ExcelPackage(fileInfo))
                {
                    // 3. Lấy trang tính đầu tiên
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    Log.Information("Worksheet Name: {SheetName}", worksheet.Name);

                    if (worksheet.Dimension == null)
                        return dataTable;

                    // 4. Lấy tổng số hàng và cột tạo cho dataTable  
                    var totalRows = worksheet.Dimension.Rows;
                    Log.Information("Total Rows in Excel ......: {TotalRows}", totalRows);

                    var totalCol = worksheet.Dimension.Columns;
                    Log.Information("Total Columns in Excel: {ColCount}", totalCol);

                    // 5. Đọc và thêm các cột GỐC từ file Excel
                    for (int colI = 1; colI <= totalCol; colI++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, colI].Value?.ToString() ?? $"Col{colI}");
                    }


                    //. Tính toán vị trí hàng bắt đầu và kết thúc của trang hiện tại
                    var startRow = (pageIndex - 1) * pageSize + 2;
                    Log.Information("Loading Page {PageIndex}, Start Row: {StartRow}", pageIndex, startRow);

                    var endRow = Math.Min(startRow + pageSize - 1, totalRows);
                    Log.Information("End Row: {EndRow}", endRow);

                    for (int row = startRow; row <= endRow; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        Log.Information("DataRow" + dataRow);

                        for (int col = 1; col <= totalCol; col++)
                        {
                            dataRow[col - 1] = worksheet.Cells[row, col].Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }

                return dataTable;
            }
            catch (Exception ex)
            {
                Log.Error("Error loading Excel page: {Message}", ex.Message);
                return default;
            }
        }

        /// <summary>
        /// Get total Record count of Excel file 
        /// </summary>
        public async Task<long> GetTotalRowCount(string filePath)
        {
            return await Task.Run(() =>
            {
                var fileInfo = new FileInfo(filePath);
                Log.Information("Getting total row count from file: {FilePath}", filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    Log.Information("Worksheet Name: {SheetName}", worksheet.Name);
                    if (worksheet.Dimension == null)
                        return 0;
                    var totalRecords = worksheet.Dimension.Rows - 1; // trừ đi 1 để loại bỏ hàng tiêu đề
                    Log.Information("Total Rows in Excel ......: {TotalRecord}", totalRecords);
                    return totalRecords;
                }
            });
        }

        /// <summary>
        /// Fuction save changes to Excel File
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="changes"></param>
        /// <returns></returns>
        public async Task SaveToFile(string filePath, List<CellChange> changes)
        {
            await Task.Run(() =>
            {
                if (filePath == null)
                    return;
                // 1. Tạo đối tượng FileInfo cho file gốc
                var fileInfo = new FileInfo(filePath);
                Log.Information("Name FilePath: {FilePath}", filePath);

                // 2. Tạo file tạm trong thư mục temp của hệ thống
                var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                Log.Information("Temp FilePath: {TempFilePath}", tempFilePath);

                //3. Tạo đối tượng FileInfo cho file tạm
                var tempFileInfo = new FileInfo(tempFilePath);
                Log.Information("Name Temp FilePath: {TempFilePath}", tempFileInfo);

                try
                {
                    //4. Copy file gốc sang file tạm
                    fileInfo.CopyTo(tempFileInfo.FullName, true);

                    //5. Mở file tạm và áp dụng các thay đổi
                    using (var package = new ExcelPackage(tempFileInfo))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        foreach (var change in changes)
                        {
                            worksheet.Cells[change.RowIndex, change.ColumnIndex].Value = change.NewValue;
                        }
                        package.Save();
                    }
                    // 6. Sau khi lưu thành công, ghi đè file gốc bằng file tạm
                    tempFileInfo.CopyTo(fileInfo.FullName, true);
                }
                catch (Exception ex)
                {
                    Log.Error("Error saving to Excel file: {Message}", ex.Message);
                }
                finally
                {

                    if (tempFileInfo.Exists)
                    {
                        tempFileInfo.Delete();
                    }
                }

            });
        }
        public async Task SaveAsToFile(DataTable dataTable, string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                MessageBox.Show("Chưa có đường dẫn để lưu , vui lòng save lại !!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await Task.Run(() =>
            {
                try
                {
                    var fi = new FileInfo(filePath);

                    if (fi.Exists)
                        fi.Delete();

                    using (var pkg = new ExcelPackage())
                    {
                        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                        // header
                        for (int c = 0; c < dataTable.Columns.Count; c++)
                            ws.Cells[1, c + 1].Value = dataTable.Columns[c].ColumnName;
                        // data
                        for (int r = 0; r < dataTable.Rows.Count; r++)
                            for (int c = 0; c < dataTable.Columns.Count; c++)
                                ws.Cells[r + 2, c + 1].Value = dataTable.Rows[r][c];
                        pkg.SaveAs(fi);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Error saving As to Excel file: {Message}", ex.Message);
                }
            });
        }
    }
}
