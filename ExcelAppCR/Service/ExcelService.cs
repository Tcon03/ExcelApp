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
using System.Security.Cryptography;
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

                        // lặp qua all các cột 
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
        public async Task SaveToFile(string filePath, List<ExcelFileInfo> changes)
        {
            await Task.Run(() =>
            {

                var fileInfo = new FileInfo(filePath);
                Log.Information("Name FilePath: {FilePath}", filePath);

                var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                Log.Information("Temp FilePath: {TempFilePath}", tempFilePath);

                var tempFileInfo = new FileInfo(tempFilePath);
                Log.Information("Name Temp FilePath: {TempFilePath}", tempFileInfo);

                // 1. Sao chép file gốc sang file tạm

                try
                {
                    // sao chép file gốc sang file tạm 
                    fileInfo.CopyTo(tempFileInfo.FullName, true);

                    // 2. Mở file tạm và áp dụng các thay đổi
                    using (var package = new ExcelPackage(tempFileInfo))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        foreach (var change in changes)
                        {
                            worksheet.Cells[change.RowIndex, change.ColumnIndex].Value = change.NewValue;
                        }
                        package.Save();
                    }
                    //ghì đè file tạm lên file gốc
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
    }
}
