using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;

namespace Test2
{
    public class ExcelService
    {
        private const int CHUNK_SIZE = 10000; // Xử lý 10K dòng mỗi lần
        private const int MAX_PREVIEW_ROWS = 1000;

        #region Public Methods

        /// <summary>
        /// Lấy dữ liệu preview từ file Excel (chỉ một số dòng đầu)
        /// </summary>
        public async Task<DataTable> GetPreviewDataAsync(string filePath, int maxRows = MAX_PREVIEW_ROWS, IProgress<ProcessingProgress> progress = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    progress?.Report(new ProcessingProgress { Message = "Opening Excel file...", Percentage = 10 });

                    var workbook = new XLWorkbook(filePath);
                    var worksheet = workbook.Worksheet(1);

                    if (worksheet.RangeUsed() == null)
                    {
                        throw new InvalidOperationException("The worksheet appears to be empty.");
                    }

                    var usedRange = worksheet.RangeUsed();
                    var totalRows = usedRange.RowCount();
                    var totalCols = usedRange.ColumnCount();

                    progress?.Report(new ProcessingProgress
                    {
                        Message = "Reading data...",
                        Percentage = 30,
                        TotalRows = totalRows
                    });

                    var dataTable = new DataTable();

                    // Đọc headers từ dòng đầu tiên
                    var headerRow = usedRange.Row(1);
                    for (int col = 1; col <= totalCols; col++)
                    {
                        var headerValue = headerRow.Cell(col).GetString();
                        var columnName = string.IsNullOrWhiteSpace(headerValue) ? $"Column{col}" : headerValue.Trim();
                        dataTable.Columns.Add(columnName, typeof(string));
                    }

                    progress?.Report(new ProcessingProgress { Message = "Reading rows...", Percentage = 50 });

                    // Đọc dữ liệu (bỏ qua dòng header)
                    var rowsToRead = Math.Min(maxRows, totalRows - 1);

                    for (int rowIndex = 2; rowIndex <= rowsToRead + 1; rowIndex++)
                    {
                        var row = usedRange.Row(rowIndex);
                        var dataRow = dataTable.NewRow();

                        for (int col = 1; col <= totalCols; col++)
                        {
                            var cellValue = row.Cell(col).GetString();
                            dataRow[col - 1] = cellValue;
                        }

                        dataTable.Rows.Add(dataRow);

                        // Báo cáo tiến độ mỗi 100 dòng
                        if (rowIndex % 100 == 0)
                        {
                            var progressPercentage = 50 + (rowIndex * 40.0 / rowsToRead);
                            progress?.Report(new ProcessingProgress
                            {
                                Message = $"Reading row {rowIndex:N0} of {rowsToRead:N0}...",
                                Percentage = progressPercentage,
                                Details = $"Loaded {dataTable.Rows.Count:N0} rows"
                            });
                        }
                    }

                    progress?.Report(new ProcessingProgress
                    {
                        Message = "Preview loaded successfully",
                        Percentage = 100,
                        Details = $"Showing {dataTable.Rows.Count:N0} of {totalRows:N0} total rows"
                    });

                    return dataTable;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Error reading Excel file: {ex.Message}", ex);
                }
            });
        }

        /// <summary>
        /// Xử lý file Excel lớn theo chunks để tối ưu bộ nhớ
        /// </summary>
        public async Task<DataTable> ProcessLargeExcelFileAsync(
            string inputPath,
            IProgress<ProcessingProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                DataTable resultTable = null;
                XLWorkbook workbook = null;
                IXLWorksheet worksheet = null;

                try
                {
                    progress?.Report(new ProcessingProgress { Message = "Opening large Excel file...", Percentage = 5 });

                    workbook = new XLWorkbook(inputPath);
                    worksheet = workbook.Worksheet(1);

                    var usedRange = worksheet.RangeUsed();
                    if (usedRange == null)
                        throw new InvalidOperationException("The worksheet appears to be empty.");

                    var totalRows = usedRange.RowCount();
                    var totalCols = usedRange.ColumnCount();

                    progress?.Report(new ProcessingProgress
                    {
                        Message = $"Processing {totalRows:N0} rows...",
                        Percentage = 10,
                        TotalRows = totalRows
                    });

                    // Tạo DataTable result
                    resultTable = CreateDataTableStructure(worksheet, totalCols);

                    // Xử lý data theo chunks
                    var processedRows = 0;
                    var startRow = 2; // Bỏ qua header

                    while (startRow <= totalRows && !cancellationToken.IsCancellationRequested)
                    {
                        var endRow = Math.Min(startRow + CHUNK_SIZE - 1, totalRows);

                        progress?.Report(new ProcessingProgress
                        {
                            Message = $"Processing rows {startRow:N0} to {endRow:N0}...",
                            Percentage = 10 + (processedRows * 80.0 / (totalRows - 1)),
                            Details = $"Processed {processedRows:N0} of {totalRows - 1:N0} rows"
                        });

                        // Xử lý chunk hiện tại
                        ProcessChunk(worksheet, resultTable, startRow, endRow, totalCols);

                        processedRows += (endRow - startRow + 1);
                        startRow = endRow + 1;

                        // Thực hiện garbage collection định kỳ để giải phóng bộ nhớ
                        if (processedRows % (CHUNK_SIZE * 5) == 0)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }

                    if (cancellationToken.IsCancellationRequested)
                    {
                        resultTable?.Dispose();
                        return null;
                    }

                    progress?.Report(new ProcessingProgress
                    {
                        Message = "Processing completed successfully",
                        Percentage = 100,
                        Details = $"Total processed: {resultTable.Rows.Count:N0} rows"
                    });

                    return resultTable;
                }
                catch (Exception ex)
                {
                    resultTable?.Dispose();
                    throw new InvalidOperationException($"Error processing large Excel file: {ex.Message}", ex);
                }
                finally
                {
                    worksheet = null;
                    workbook?.Dispose();
                }
            }, cancellationToken);
        }

        /// <summary>
        /// Lưu DataTable ra file Excel với tối ưu hiệu năng
        /// </summary>
        public async Task SaveToExcelAsync(
            DataTable dataTable,
            string outputPath,
            IProgress<ProcessingProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            await Task.Run(() =>
            {
                try
                {
                    progress?.Report(new ProcessingProgress { Message = "Creating Excel workbook...", Percentage = 10 });

                     var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("ProcessedData");

                    progress?.Report(new ProcessingProgress { Message = "Writing headers...", Percentage = 20 });

                    // Viết headers
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
                        worksheet.Cell(1, col + 1).Style.Font.Bold = true;
                        worksheet.Cell(1, col + 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    }

                    progress?.Report(new ProcessingProgress { Message = "Writing data rows...", Percentage = 30 });

                    // Viết dữ liệu theo chunks
                    var totalRows = dataTable.Rows.Count;
                    var writtenRows = 0;

                    for (int startIdx = 0; startIdx < totalRows; startIdx += CHUNK_SIZE)
                    {
                        if (cancellationToken.IsCancellationRequested) return;

                        var endIdx = Math.Min(startIdx + CHUNK_SIZE, totalRows);

                        for (int rowIdx = startIdx; rowIdx < endIdx; rowIdx++)
                        {
                            var dataRow = dataTable.Rows[rowIdx];
                            var excelRowIdx = rowIdx + 2; // +1 for Excel 1-based, +1 for header

                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                var value = dataRow[col];
                                worksheet.Cell(excelRowIdx, col + 1).Value = value?.ToString() ?? "";
                            }
                        }

                        writtenRows = endIdx;
                        var progressPercentage = 30 + (writtenRows * 60.0 / totalRows);

                        progress?.Report(new ProcessingProgress
                        {
                            Message = $"Writing row {writtenRows:N0} of {totalRows:N0}...",
                            Percentage = progressPercentage,
                            Details = $"Saved {writtenRows:N0} rows"
                        });
                    }

                    if (cancellationToken.IsCancellationRequested) return;

                    progress?.Report(new ProcessingProgress { Message = "Formatting and saving...", Percentage = 95 });

                    // Auto-fit columns (chỉ cho vài cột đầu để tránh chậm)
                    var colsToAutoFit = Math.Min(10, dataTable.Columns.Count);
                    for (int col = 1; col <= colsToAutoFit; col++)
                    {
                        worksheet.Column(col).AdjustToContents();
                    }

                    // Lưu file
                    workbook.SaveAs(outputPath);

                    progress?.Report(new ProcessingProgress
                    {
                        Message = "File saved successfully",
                        Percentage = 100,
                        Details = $"Saved {totalRows:N0} rows to {Path.GetFileName(outputPath)}"
                    });
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Error saving Excel file: {ex.Message}", ex);
                }
            }, cancellationToken);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Tạo cấu trúc DataTable dựa trên worksheet
        /// </summary>
        private DataTable CreateDataTableStructure(IXLWorksheet worksheet, int totalCols)
        {
            var dataTable = new DataTable();
            var headerRow = worksheet.Row(1);

            for (int col = 1; col <= totalCols; col++)
            {
                var headerValue = headerRow.Cell(col).GetString();
                var columnName = string.IsNullOrWhiteSpace(headerValue) ? $"Column{col}" : headerValue.Trim();

                // Đảm bảo tên cột là duy nhất
                var originalName = columnName;
                var counter = 1;
                while (dataTable.Columns.Contains(columnName))
                {
                    columnName = $"{originalName}_{counter++}";
                }

                dataTable.Columns.Add(columnName, typeof(string));
            }

            return dataTable;
        }

        /// <summary>
        /// Xử lý một chunk dữ liệu và thực hiện business logic
        /// </summary>
        private void ProcessChunk(IXLWorksheet worksheet, DataTable resultTable, int startRow, int endRow, int totalCols)
        {
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                var row = worksheet.Row(rowIndex);
                var dataRow = resultTable.NewRow();

                // Đọc dữ liệu từ Excel
                for (int col = 1; col <= totalCols; col++)
                {
                    var cellValue = row.Cell(col).GetString();
                    dataRow[col - 1] = cellValue;
                }

                // ===== BUSINESS LOGIC XỬ LÝ DỮ LIỆU =====
                // Đây là nơi bạn thực hiện các xử lý dữ liệu cụ thể
                dataRow = ProcessBusinessLogic(dataRow, resultTable.Columns.Count);

                resultTable.Rows.Add(dataRow);
            }
        }

        /// <summary>
        /// Thực hiện business logic xử lý dữ liệu
        /// Bạn có thể customize method này theo yêu cầu cụ thể
        /// </summary>
        private DataRow ProcessBusinessLogic(DataRow inputRow, int columnCount)
        {
            // ===== VÍ DỤ CÁC XỬ LÝ DỮ LIỆU =====

            // 1. Làm sạch dữ liệu
            for (int i = 0; i < columnCount; i++)
            {
                var value = inputRow[i]?.ToString()?.Trim();
                inputRow[i] = string.IsNullOrEmpty(value) ? "" : value;
            }

            // 2. Validate và transform dữ liệu
            // Ví dụ: Chuẩn hóa số điện thoại, email, v.v.
            if (columnCount > 0) // Giả sử cột 0 là tên
            {
                var name = inputRow[0]?.ToString();
                if (!string.IsNullOrEmpty(name))
                {
                    // Chuẩn hóa tên: Title Case
                    inputRow[0] = ToTitleCase(name);
                }
            }

            if (columnCount > 1) // Giả sử cột 1 là email
            {
                var email = inputRow[1]?.ToString();
                if (!string.IsNullOrEmpty(email))
                {
                    inputRow[1] = email.ToLower().Trim();
                }
            }

            // 3. Thêm các cột tính toán mới (nếu cần)
            // Ví dụ: Thêm timestamp xử lý
            // inputRow["ProcessedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // 4. Business rules validation
            // Ví dụ: Đánh dấu các record không hợp lệ
            // if (SomeValidationRule(inputRow))
            // {
            //     inputRow["IsValid"] = "Yes";
            // }

            return inputRow;
        }

        /// <summary>
        /// Chuyển đổi string thành Title Case
        /// </summary>
        private string ToTitleCase(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var words = input.Split(' ', (char)StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 0)
                {
                    words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1).ToLower();
                }
            }
            return string.Join(" ", words);
        }

        #endregion
    }

    #region Supporting Classes

    /// <summary>
    /// Class để báo cáo tiến độ xử lý
    /// </summary>
    public class ProcessingProgress
    {
        public string Message { get; set; } = "";
        public double Percentage { get; set; } = 0;
        public string Details { get; set; } = "";
        public long TotalRows { get; set; } = 0;
        public long ProcessedRows { get; set; } = 0;
    }

    #endregion
}
