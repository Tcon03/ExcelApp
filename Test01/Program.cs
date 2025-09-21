using OfficeOpenXml;
using System.Data;
using System; 
using System.IO; 
using Newtonsoft.Json;
namespace Test01
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Nguyen Truong");
            

            var filePath = "DuLieuExcel.xlsx";

            Console.WriteLine("--- Bắt đầu quá trình GHI file Excel ---");
            WriteToExcel(filePath);

            Console.WriteLine($"Đã tạo và ghi dữ liệu thành công vào file: {Path.GetFullPath(filePath)}\n");

            Console.WriteLine("--- Bắt đầu quá trình ĐỌC file Excel ---");
            Console.WriteLine("Vui lòng tạo file 'TestDataRead.xlsx' với dữ liệu mẫu để chương trình có thể đọc.");

            // Kiểm tra xem file đọc có tồn tại không
            if (File.Exists(filePath))
            {
                var userDetails = ReadFromExcel<List<UserDetails>>(filePath);
                Console.WriteLine("User details :" + userDetails.ToString());
                Console.WriteLine("Dữ liệu đọc được từ file Excel:");
                foreach (var user in userDetails)
                {
                    Console.WriteLine(user.ToString());
                }
            }
            else
            {
                Console.WriteLine($"Không tìm thấy file '{filePath}'. Vui lòng tạo file này theo hướng dẫn ở Bước 5.");
            }

            Console.WriteLine("\nChương trình kết thúc. Nhấn phím bất kỳ để thoát.");
            Console.ReadKey();
        }

        private static void WriteToExcel(string path)
        {
            // Dữ liệu mẫu để ghi vào file
            List<UserDetails> persons = new List<UserDetails>()
        {
            new UserDetails() { ID = "9999", Name = "ABCD", City = "City1", Country = "USA" },
            new UserDetails() { ID = "8888", Name = "PQRS", City = "City2", Country = "INDIA" },
            new UserDetails() { ID = "7777", Name = "XYZZ", City = "City3", Country = "CHINA" },
            new UserDetails() { ID = "6666", Name = "LMNO", City = "City4", Country = "UK" },
        };

            // Chuyển danh sách đối tượng thành DataTable để dễ xử lý
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));

            if(File.Exists(path))
            {
                File.Delete(path);
            }    
            FileInfo filePath = new FileInfo(path); 
            using (var excelPack = new ExcelPackage(filePath))
            {
                var ws = excelPack.Workbook.Worksheets.Add("WriteTest");
                // Ghi dữ liệu từ DataTable vào worksheet, bắt đầu từ ô A1
                // tham số 'true' để ghi cả tên cột (header)
                ws.Cells["A1"].LoadFromDataTable(table, true);
                excelPack.Save();
            }
        }

        private static T ReadFromExcel<T>(string path, bool hasHeader = true)
        {
           
            using (var excelPack = new ExcelPackage())
            {
                //Load excel stream
                using (var stream = File.OpenRead(path))
                {
                    excelPack.Load(stream);
                }

                //Lets Deal with first worksheet.(You may iterate here if dealing with multiple sheets)
                var ws = excelPack.Workbook.Worksheets[0];
                Console.WriteLine("ws :" +ws );
                //Get all details as DataTable -because Datatable make life easy :)
                DataTable excelasTable = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    //Get colummn details
                    if (!string.IsNullOrEmpty(firstRowCell.Text))
                    {
                        string firstColumn = string.Format("Column {0}", firstRowCell.Start.Column);
                        Console.WriteLine("first Column" + firstColumn);
                        excelasTable.Columns.Add(hasHeader ? firstRowCell.Text : firstColumn);
                    }
                }
                var startRow = hasHeader ? 2 : 1;
                Console.WriteLine("start Row :" + startRow );
                //Get row details
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, excelasTable.Columns.Count];
                    DataRow row = excelasTable.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                //Get everything as generics and let end user decides on casting to required type
                var generatedType = JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(excelasTable));

                return (T)Convert.ChangeType(generatedType, typeof(T));
            }
        }
    }


}
