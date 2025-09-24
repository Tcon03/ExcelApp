using ExcelAppCR.Model;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAppCR.Service
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
        /// <param name="filePath"> Đường dẫn của file Excel </param>
        /// <param name="pageIndex"> số thứ tự page muốn lấy </param>
        /// <param name="pageSize"> Kích thước cuẩ mỗi page 100 ,1000 </param>
        /// <returns> dataTable chứa dữ liệu</returns>
        public DataTable LoadExcelPage(string filePath, int pageIndex, int pageSize)
        {
            try
            {
                // 1. lấy ra một dataTable rỗng
                var dataTable = new DataTable();

                //2. xác định vị trí file excel
                var fileInfo = new FileInfo(filePath);
                Log.Information("Getting total row count from file: {FilePath}", filePath);

                //3 . mở file excel và đọc dữ liệu
                using (var package = new ExcelPackage(fileInfo))
                {
                    // 3.1 mở file excel và lấy ra sheet đầu tiên
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    Log.Information("Worksheet Name: {SheetName}", worksheet.Name);

                    //3.2 kiểm tra file excel có dữ liệu hay không nếu kh thì trả về dataTable rỗng
                    if (worksheet.Dimension == null)
                        return dataTable;

                    //3.3 đếm xem có bao nhiêu dòng trong file excel
                    var totalRows = worksheet.Dimension.Rows;
                    Log.Information("Total Rows in Excel ......: {TotalRows}", totalRows);

                    //3.4 Đếm xem có bao nhiêu cột trong file excel
                    var totalCol = worksheet.Dimension.Columns;
                    Log.Information("Total Columns in Excel: {ColCount}", totalCol);


                    // 3.5 nhìn vào cột đầu tiên và tạo các column tương ứng trong dataTable
                    for (int colI = 1; colI <= totalCol; colI++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, colI].Value?.ToString() ?? $"Col{colI}");
                    }

                    // Vì dòng 1 của Excel là dòng tiêu đề, nên dữ liệu thực tế bắt đầu từ dòng 2.Vậy nên, dòng dữ liệu thứ 10 sẽ
                    //nằm ở hàng 11 trong Excel. Dòng bắt đầu của trang 2 sẽ là hàng 12.Phép toán(2 - 1) * 10 + 2 = 12 là hoàn toàn chính xác.
                    var startRow = (pageIndex - 1) * pageSize + 2;
                    Log.Information("Loading Page {PageIndex}, Start Row: {StartRow}", pageIndex, startRow);

                    //hàng kết thúc được lấy ra là start + sizeof - 1 header 
                    var endRow = Math.Min(startRow + pageSize - 1, totalRows);
                    Log.Information("End Row: {EndRow}", endRow);

                    // Chỉ lặp qua thời điểm bắt đầu và đk kết thúc
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        Log.Information("Reading Row: {Row}", row);

                        // lắp qua all các cột 
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
        public async Task<long> GetTotalRowCount(string filePath)
        {
            var fileInfo = new FileInfo(filePath);
            Log.Information("Getting total row count from file: {FilePath}", filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                Log.Information("Worksheet Name: {SheetName}", worksheet.Name);
                if (worksheet.Dimension == null)
                    return 0;
                var totalRows = worksheet.Dimension.Rows - 1; // trừ đi 1 để loại bỏ hàng tiêu đề
                Log.Information("Total Rows in Excel ......: {TotalRows}", totalRows);
                return totalRows;
            }
        }
        public async Task SaveToFile(DataTable dataTable, string filePath)
        {
            await Task.Run(() =>
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);

                // Tạo file tạm để backup
                string tempFile = filePath + ".temp";

                // Copy file gốc sang file tạm
                File.Copy(filePath, tempFile, true);

                try
                {

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

                catch (Exception ex)
                {
                    Log.Error("Error saving to Excel file: {Message}", ex.Message);
                    // Nếu có lỗi xảy ra, khôi phục lại file gốc từ file tạm
                    if (File.Exists(tempFile))
                    {
                        File.Copy(tempFile, filePath, true);
                    }

                }
                finally
                {
                    // Xóa file tạm sau khi hoàn thành (bất kể thành công hay thất bại)
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }

            });
        }
    }
}
