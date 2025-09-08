using ClosedXML.Excel;
using System.IO;
using System.Windows;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace Viettel_Report_Automation.Controllers
{
    public class WordReportController
    {

        public void generateWordFile(IProgress<string> progress, string fileTheoDoi = "", string sheetTheoDoi = "")
        {
            progress.Report("Đang tạo báo cáo word");

            var wb = new XLWorkbook(fileTheoDoi);
            var ws = wb.Worksheet("BC_chi_tiet");

            var wsMeta = wb.Worksheet("Meta");
            List<string> mainCreatia = new List<string>();

            string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "word_mau.docx");

            using (var doc = DocX.Load(wordPath))
            {
                string font = "Times New Roman";
                doc.ReplaceText("{thang}", "08");
                doc.ReplaceText("{nam}", "2025");

                doc.ReplaceText("{nhanxet01}", "Vị trí trạm hiện tại Viettel đang chiếm ưu thế với 1633 vị trí. Số lượng vị trí trạm Viettel nhiều hơn Vinaphone 240 vị trí và nhiều hơn Mobifone 510 vị trí. Xét về mức huyện Viettel còn 4 huyện có vị trí trạm ít hơn nhà mạng Vina là Krông Bông ít hơn 4 vị trí, Huyện Ea Súp và Krông Búk ít hơn 1 vị trí, huyện Ea Súp ít hơn 5 trạm");
                doc.ReplaceText("{h_ketquathuchien6thang}", "KẾT QUẢ THỰC HIỆN 6 THÁNG ĐẦU NĂM 2025");

                // tao bang
                var table = doc.AddTable(2, 11);
                table.Design = TableDesign.TableGrid;
                table.MergeCellsInColumn(0, 0, 1);
                table.MergeCellsInColumn(1, 0, 1);
                table.MergeCellsInColumn(2, 0, 1);
                table.MergeCellsInColumn(3, 0, 1);
                table.MergeCellsInColumn(4, 0, 1);
                table.Rows[0].MergeCells(5, 7);
                table.Rows[0].MergeCells(6, 8);

                table.Rows[0].Cells[0].Paragraphs[0].Append(@"Các chỉ tiêu triển khai hạ tầng chính").Font(font).Bold();
                table.Rows[0].Cells[1].Paragraphs[0].Append(@"ĐVT").Font(font).Bold();
                table.Rows[0].Cells[2].Paragraphs[0].Append(@"Kế hoạch T7").Font(font).Bold();
                table.Rows[0].Cells[3].Paragraphs[0].Append(@"Thực hiện T7").Font(font).Bold();
                table.Rows[0].Cells[4].Paragraphs[0].Append(@"%TH").Font(font).Bold();
                table.Rows[0].Cells[5].Paragraphs[0].Append(@"Thực hiện năm 2025").Font(font).Bold();
                table.Rows[0].Cells[6].Paragraphs[0].Append(@"Thực hiện năm 2024").Font(font).Bold();

                table.Rows[1].Cells[5].Paragraphs[0].Append(@"Kế hoạch").Font(font).Bold();
                table.Rows[1].Cells[6].Paragraphs[0].Append(@"Thực hiện").Font(font).Bold();
                table.Rows[1].Cells[7].Paragraphs[0].Append(@"%TH").Font(font).Bold();

                table.Rows[1].Cells[8].Paragraphs[0].Append(@"Kế hoạch").Font(font).Bold();
                table.Rows[1].Cells[9].Paragraphs[0].Append(@"Thực hiện").Font(font).Bold();
                table.Rows[1].Cells[10].Paragraphs[0].Append(@"%TH").Font(font).Bold();
                /*-----------------------------------------------------------------------*/
                table.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[3].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[4].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[5].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[6].Paragraphs[0].Alignment = Alignment.center;

                table.Rows[1].Cells[5].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[6].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[7].Paragraphs[0].Alignment = Alignment.center;

                table.Rows[1].Cells[8].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[9].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[10].Paragraphs[0].Alignment = Alignment.center;

                /*-----------------------------------------------------------------------*/
               
                int rowTable = 2;
                for (int row = 3; row < ws.RowsUsed().Count(); row++)
                {
                    table.InsertRow();
                    string dauMuc = ws.Cell($"B{row}").Value.ToString();
                    table.Rows[rowTable].Cells[0].Paragraphs[0].Append(dauMuc).Font(font);
                    table.Rows[rowTable].Cells[1].Paragraphs[0].Append(ws.Cell($"C{row}").Value.ToString()).Font(font);
                    
                    table.Rows[rowTable].Cells[2].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"G{row}").FormulaA1, wsMeta)).Font(font);
                    table.Rows[rowTable].Cells[3].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"H{row}").FormulaA1, wsMeta)).Font(font);
                    
                    table.Rows[rowTable].Cells[4].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"I{row}").FormulaA1, wsMeta)).Font(font);
                    table.Rows[rowTable].Cells[5].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"M{row}").FormulaA1, wsMeta)).Font(font);
                    table.Rows[rowTable].Cells[6].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"N{row}").FormulaA1, wsMeta)).Font(font);
                    table.Rows[rowTable].Cells[7].Paragraphs[0].Append(getValueFromFormula(ws.Cell($"O{row}").FormulaA1, wsMeta)).Font(font);
                    //table.Rows[row].Cells[8].Paragraphs[0].Append(ws.Cell($"B{row}").Value.ToString()).Font(font);
                    //table.Rows[row].Cells[9].Paragraphs[0].Append(ws.Cell($"B{row}").Value.ToString()).Font(font);
                    //table.Rows[row].Cells[10].Paragraphs[0].Append(ws.Cell($"B{row}").Value.ToString()).Font(font);
                    rowTable++;
                }

                var p = doc.Paragraphs.Where(s => s.Text.Contains("{bangchitieu}")).FirstOrDefault();
                p.ReplaceText("{bangchitieu}", "");
                p.InsertTableAfterSelf(table);

                // Bang trien khai xay dung tram BTS moi
                /*  var tableBTS = doc.AddTable(5, 9);
                  tableBTS.Design = TableDesign.TableGrid;
                  tableBTS.MergeCellsInColumn(0, 0, 1);
                  tableBTS.MergeCellsInColumn(1, 0, 1);
                  tableBTS.MergeCellsInColumn(2, 0, 1);
                  tableBTS.Rows[0].MergeCells(3, 5);
                  tableBTS.Rows[0].MergeCells(4, 6);
                  tableBTS.Rows[0].Cells[0].Paragraphs[0].Append(@"STT").Font(font).Bold();
                  tableBTS.Rows[0].Cells[1].Paragraphs[0].Append(@"Đối tác XHH").Font(font).Bold();
                  tableBTS.Rows[0].Cells[2].Paragraphs[0].Append(@"Quỹ trạm năm 2025").Font(font).Bold();
                  tableBTS.Rows[0].Cells[3].Paragraphs[0].Append(@"Tiến độ hiện tại").Font(font).Bold();
                  tableBTS.Rows[0].Cells[4].Paragraphs[0].Append(@"Còn lại thực hiện").Font(font).Bold();

                  tableBTS.Rows[1].Cells[3].Paragraphs[0].Append(@"Thuê").Font(font).Bold();
                  tableBTS.Rows[1].Cells[4].Paragraphs[0].Append(@"Khởi công").Font(font).Bold();
                  tableBTS.Rows[1].Cells[5].Paragraphs[0].Append(@"ĐBHT").Font(font).Bold();
                  tableBTS.Rows[1].Cells[6].Paragraphs[0].Append(@"Thuê").Font(font).Bold();
                  tableBTS.Rows[1].Cells[7].Paragraphs[0].Append(@"Khởi công").Font(font).Bold();
                  tableBTS.Rows[1].Cells[8].Paragraphs[0].Append(@"ĐBHT").Font(font).Bold();

                  tableBTS.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[0].Cells[3].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[0].Cells[4].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[3].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[4].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[5].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[6].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[7].Paragraphs[0].Alignment = Alignment.center;
                  tableBTS.Rows[1].Cells[8].Paragraphs[0].Alignment = Alignment.center;

                  tableBTS.Rows[2].Cells[0].Paragraphs[0].Append("1");
                  tableBTS.Rows[3].Cells[0].Paragraphs[0].Append("2");

                  tableBTS.Rows[2].Cells[1].Paragraphs[0].Append("DLK");
                  tableBTS.Rows[3].Cells[1].Paragraphs[0].Append("PYN");

                  tableBTS.Rows[2].Cells[2].Paragraphs[0].Append("20");
                  tableBTS.Rows[3].Cells[2].Paragraphs[0].Append("88");

                  tableBTS.Rows[4].Cells[1].Paragraphs[0].Append("Tổng");
                  p = doc.Paragraphs.Where(s => s.Text.Contains("{t_trienkhaixaydungbtsmoi}")).First();
                  p.ReplaceText("{t_trienkhaixaydungbtsmoi}", "");
                  p.InsertTableAfterSelf(tableBTS);
                */

                wb.Dispose();
                doc.Save();
            }
            progress.Report("Đã tạo file word");
        }

        private string getValueFromFormula(string formula, IXLWorksheet worksheet)
        {
            if (formula.Contains("Meta"))
            {
                string s = formula.Replace("=", "");
                string[] sp = s.Split("!");
                return worksheet.Cell(sp[1]).Value.ToString();
            }
            return "0";
        }
        
    }
}
