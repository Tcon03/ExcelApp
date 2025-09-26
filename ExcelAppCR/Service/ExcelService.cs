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
        /// <param name="pageSize"> Kích thước của mỗi page 100 ,1000 </param>
        /// <returns> dataTable chứa dữ liệu</returns>
        public DataTable LoadExcelPage(string filePath, int pageIndex, int pageSize)
        {
            try
            {
                //1. Khởi tạo
                var dataTable = new DataTable();

                var fileInfo = new FileInfo(filePath);
                Log.Information("Getting total row count from file: {FilePath}", filePath);

                //2 . mở file excel và đọc dữ liệu
                using (var package = new ExcelPackage(fileInfo))
                {
                    //3. Lấy trang tính đầu tiên
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    Log.Information("Worksheet Name: {SheetName}", worksheet.Name);

                    if (worksheet.Dimension == null)
                        return dataTable;

                    //4. Lấy tổng số hàng và cột tạo cho dataTable  
                    var totalRows = worksheet.Dimension.Rows;
                    Log.Information("Total Rows in Excel ......: {TotalRows}", totalRows);

                    var totalCol = worksheet.Dimension.Columns;
                    Log.Information("Total Columns in Excel: {ColCount}", totalCol);


                    for (int colI = 1; colI <= totalCol; colI++)
                    {
                        dataTable.Columns.Add(worksheet.Cells[1, colI].Value?.ToString() ?? $"Col{colI}");
                    }

                    //5. Tính toán vị trí hàng bắt đầu và kết thúc của trang hiện tại
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
        }

        /// <summary>
        /// Fuction save changes to Excel File
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="changes"></param>
        /// <returns></returns>
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
                    // copy đè file tạm lên file gốc
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
