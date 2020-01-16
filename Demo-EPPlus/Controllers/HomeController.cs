using Demo_EPPlus.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Demo_EPPlus.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        private List<TestItemClass> CreateTestItems()
        {
            var resultsList = new List<TestItemClass>();

            for (int i = 0; i < 15; i++)
            {
                var a = new TestItemClass()
                {
                    Id = i,
                    Address = "Test Excel Address at " + i,
                    Money = 20000 + i * 10,
                    FullName = "TMSANG " + i
                };
                resultsList.Add(a);
            }

            return resultsList;
        }

        private Stream CreateExcelFile(Stream stream = null)
        {
            var list = CreateTestItems();
            using (var excelPakage = new ExcelPackage(stream ?? new MemoryStream()))
            {
                // Tạo author cho file excel
                excelPakage.Workbook.Properties.Author = "Hanker";

                // Tạo title
                excelPakage.Workbook.Properties.Title = "EPP test background";

                // Thêm comment
                excelPakage.Workbook.Properties.Comments = "This is my fucking generated comment";

                // Thêm sheet
                excelPakage.Workbook.Worksheets.Add("First sheet");

                // Lấy sheet vừa thêm để thao tác
                var worksheet = excelPakage.Workbook.Worksheets[1];

                // Đổ data vào file excel
                worksheet.Cells[1, 1].LoadFromCollection(list, true);
                BindingFormatForExcel(worksheet, list);
                excelPakage.Save();
                return excelPakage.Stream;
            }
        }

        private void BindingFormatForExcel(ExcelWorksheet worksheet, List<TestItemClass> listItems)
        {
            // Set default width cho tất cả column
            worksheet.DefaultColWidth = 10;

            // Tự động xuống hàng khi text quá dài
            worksheet.Cells.Style.WrapText = true;

            // Tạo header
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Full Name";
            worksheet.Cells[1, 3].Value = "Address";
            worksheet.Cells[1, 4].Value = "Money";

            // Lấy range vào tạo format cho range đó. Ở đây là từ A1 tới D1
            using (var range = worksheet.Cells["A1:D1"])
            {
                // Set PatternType
                range.Style.Fill.PatternType = ExcelFillStyle.DarkGray;

                // Set Màu cho Background
                range.Style.Fill.BackgroundColor.SetColor(Color.Aqua);

                // Canh giữa cho các text
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Set font cho text tong range hiện tại
                range.Style.Font.SetFromFont(new Font("Arial", 10));

                // Set border
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                // Set màu cho border
                range.Style.Border.Bottom.Color.SetColor(Color.Blue);
            }

            // Đổ dữ liệu từ list vào
            for (int i = 0; i < listItems.Count; i++)
            {
                var item = listItems[i];
                worksheet.Cells[i + 2, 1].Value = item.Id + 1;
                worksheet.Cells[i + 2, 2].Value = item.FullName;
                worksheet.Cells[i + 2, 3].Value = item.Address;
                worksheet.Cells[i + 2, 4].Value = item.Money;

                // Format lại color nếu như thỏa điều kiện
                if (item.Money > 20050)
                {
                    using (var range = worksheet.Cells[i + 2, 1, i + 2, 4])
                    {
                        range.Style.Font.Color.SetColor(Color.Red);
                        range.Style.Font.Bold = true;
                    }
                }
            }

            // Canh giữa cho cột Id
            worksheet.Cells[1, 1, listItems.Count + 1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // Format lại định dạng xuất ra ở cột Money
            worksheet.Cells[2, 4, listItems.Count + 4, 4].Style.Numberformat.Format = "$#,##.00";

            // Fix lại width của column với minium width là 15
            worksheet.Cells[1, 1, listItems.Count + 5, 1].AutoFitColumns(5);
            worksheet.Cells[1, 2, listItems.Count + 5, 2].AutoFitColumns(15);
            worksheet.Cells[1, 3, listItems.Count + 5, 3].AutoFitColumns(30);
            worksheet.Cells[1, 4, listItems.Count + 5, 4].AutoFitColumns(15);

            // Thực hiện tính theo formula trong excel
            // Hàm Sum
            worksheet.Cells[listItems.Count + 3, 3].Value = "Total is: ";
            worksheet.Cells[listItems.Count + 3, 4].Formula = String.Format("SUM(D2:D{0})", listItems.Count + 1);

            // Hàm SumIf
            worksheet.Cells[listItems.Count + 4, 3].Value = "Greater than 20050: ";
            worksheet.Cells[listItems.Count + 4, 4].Formula = String.Format("SUMIF(D2:D{0}, \">20050\")", listItems.Count + 1);

            // Tính theo %
            worksheet.Cells[listItems.Count + 5, 3].Value = "Percentage: ";
            worksheet.Cells[listItems.Count + 5, 4].Style.Numberformat.Format = "0.00%";
            worksheet.Cells[listItems.Count + 5, 4].FormulaR1C1 = "(R[-1]C/R[-2]C)";
        }

        [HttpGet]
        public ActionResult Export()
        {
            // Gọi lại hàm để tạo file excel
            var stream = CreateExcelFile();

            // Tạo buffer memory stream để hứng file excel
            var buffer = stream as MemoryStream;

            // Đây là context type dành cho file excel, còn rất nhiều content type khác
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            // Dòng này quan trọng vì chạy trên firefox hay IE thì dòng này sẽ hiện Save As dialog cho người dùng chọn thư mục để lưu
            // File name là ExcelDemo.xlsx
            Response.AddHeader("Content-Disposition", "attachment; filename=ExcelDemo.xlsx");

            // Lưu file excel như một mảng byte để trả về response
            Response.BinaryWrite(buffer.ToArray());

            // Send tất cả output bytes về phía clients
            Response.Flush();
            Response.End();

            // Redirect về luôn trang index
            return RedirectToAction("Index");
        }

        private DataTable ReadFromExcelFile(string path, string sheetName)
        {
            // Khởi tạo data table
            DataTable dt = new DataTable();

            // Load file excel và các setting ban đầu
            using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
            {
                if (package.Workbook.Worksheets.Count < 1)
                {
                    // Log - Không có sheet nào tồn tại trong file excel của bạn
                    return null;
                }
                // Lấy sheet đầu tiên trong file Excel để truy vấn, truyền vào name của sheet để lấy ra sheet cần, 
                // nếu name = null thì lấy sheet đầu tiên
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName) ?? package.Workbook.Worksheets.FirstOrDefault();

                // Đọc tất cả các header
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dt.Columns.Add(firstRowCell.Text);
                }

                // Đọc tất cả data bắt đầu từ row thứ 2
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    // Lấy 1 row trong excel để truy vấn
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];

                    // Tạo 1 row trong data table
                    var newRow = dt.NewRow();
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }
                    dt.Rows.Add(newRow);
                }
            }
            return dt;
        }

        [HttpGet]
        public ActionResult ReadFromExcel()
        {
            var data = ReadFromExcelFile(@"D:\ExcelDemo.xlsx", "First sheet");
            return View(data);
        }
    }
}