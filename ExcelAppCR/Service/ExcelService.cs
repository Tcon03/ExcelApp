using ExcelAppCR.Model;
using OfficeOpenXml;
using OfficeOpenXml.Table;
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
        /// Load dữ liệu từ file Excel
        /// </summary>
        /// <param name="filePath">Đường dẫn file Excel</param>
        /// <returns>DataTable chứa dữ liệu</returns>
        public async Task<DataTable> LoadExcelDataAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                var dataTable = new DataTable();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                        throw new InvalidOperationException("File Excel không có worksheet nào.");

                    var worksheet = package.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null)
                        return dataTable; // Trả về DataTable rỗng nếu không có dữ liệu

                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;

                    // Tạo columns từ hàng đầu tiên (header)
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        var headerValue = worksheet.Cells[start.Row, col].Value?.ToString();
                        if (string.IsNullOrWhiteSpace(headerValue))
                            headerValue = $"Column{col}";

                        dataTable.Columns.Add(headerValue);
                    }

                    // Đọc dữ liệu từ hàng thứ 2
                    for (int row = start.Row + 1; row <= end.Row; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        bool hasData = false;

                        for (int col = start.Column; col <= end.Column; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            dataRow[col - start.Column] = cellValue?.ToString() ?? string.Empty;

                            if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                                hasData = true;
                        }

                        // Chỉ thêm row nếu có dữ liệu
                        if (hasData)
                            dataTable.Rows.Add(dataRow);
                    }
                }
                return dataTable;
            });
        }

        /// <summary>
        /// Lấy thông tin cơ bản về file Excel
        /// </summary>
        /// <param name="filePath">Đường dẫn file Excel</param>
        /// <returns>Thông tin file</returns>
        public ExcelFileInfo GetExcelInfo(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    return new ExcelFileInfo { SheetCount = 0, RowCount = 0, ColumnCount = 0 };

                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                    return new ExcelFileInfo { SheetCount = package.Workbook.Worksheets.Count, RowCount = 0, ColumnCount = 0 };

                return new ExcelFileInfo
                {
                    SheetCount = package.Workbook.Worksheets.Count,
                    RowCount = worksheet.Dimension.End.Row,
                    ColumnCount = worksheet.Dimension.End.Column,
                    SheetName = worksheet.Name
                };
            }
        }
    }
}
