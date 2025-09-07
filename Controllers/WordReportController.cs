using ClosedXML.Excel;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Viettel_Report_Automation.Controllers
{
    public class WordReportController
    {

        public void generateWordFile(IProgress<string> progress, string fileTheoDoi ="")
        {

            var wb = new XLWorkbook(fileTheoDoi);
            var ws = wb.Worksheet("meta");
            List<string> mainCreatia = new List<string>();

            string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "word_mau.docx");

            using (var doc = DocX.Load(wordPath))
            {
                string font = "Times New Roman";
                doc.ReplaceText("{thang}", "08");
                doc.ReplaceText("{nam}", "2025");

                doc.ReplaceText("{nhanxet01}", "Vị trí trạm hiện tại Viettel đang chiếm ưu thế với 1633 vị trí. Số lượng vị trí trạm Viettel nhiều hơn Vinaphone 240 vị trí và nhiều hơn Mobifone 510 vị trí. Xét về mức huyện Viettel còn 4 huyện có vị trí trạm ít hơn nhà mạng Vina là Krông Bông ít hơn 4 vị trí, Huyện Ea Súp và Krông Búk ít hơn 1 vị trí, huyện Ea Súp ít hơn 5 trạm");


                // tao bang
                var table = doc.AddTable(11, 11);
                table.Design = TableDesign.TableGrid;
                table.MergeCellsInColumn(0, 0, 1);
                table.MergeCellsInColumn(1, 0, 1);
                table.MergeCellsInColumn(2, 0, 1);
                table.MergeCellsInColumn(3, 0, 1);
                table.MergeCellsInColumn(4, 0, 1);
                table.Rows[0].MergeCells(5, 7);
                table.Rows[0].MergeCells(6, 8);
                

                table.Rows[0].Cells[0].Paragraphs[0].Append(@"Các chỉ tiêu triển khai hạ tầng chính").Font(font);
                table.Rows[0].Cells[1].Paragraphs[0].Append(@"ĐVT").Font(font);
                table.Rows[0].Cells[2].Paragraphs[0].Append(@"Kế hoạch T7").Font(font);
                table.Rows[0].Cells[3].Paragraphs[0].Append(@"Thực hiện T7").Font(font);
                table.Rows[0].Cells[4].Paragraphs[0].Append(@"%TH").Font(font);
                table.Rows[0].Cells[5].Paragraphs[0].Append(@"Thực hiện năm 2025").Font(font);
                table.Rows[0].Cells[6].Paragraphs[0].Append(@"Thực hiện năm 2024").Font(font);

                table.Rows[1].Cells[5].Paragraphs[0].Append(@"Kế hoạch").Font(font);
                table.Rows[1].Cells[6].Paragraphs[0].Append(@"Thực hiện").Font(font);
                table.Rows[1].Cells[7].Paragraphs[0].Append(@"%TH").Font(font);

                table.Rows[1].Cells[8].Paragraphs[0].Append(@"Kế hoạch").Font(font);
                table.Rows[1].Cells[9].Paragraphs[0].Append(@"Thực hiện").Font(font);
                table.Rows[1].Cells[10].Paragraphs[0].Append(@"%TH").Font(font);

                var p = doc.Paragraphs.Where(s => s.Text.Contains("{bangchitieu}")).FirstOrDefault();
                p.ReplaceText("{bangchitieu}", "");
                p.InsertTableAfterSelf(table);

                doc.Save();
            }
                progress.Report("Đã tạo file word");
        }
    }
}
