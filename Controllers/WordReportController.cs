using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore.Storage;
using System.Data;
using System.Drawing;
using System.IO;
using Viettel_Report_Automation.Utils;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Viettel_Report_Automation.Controllers
{
    public class WordReportController
    {
        private string font = "Times New Roman";
        private void hatangdidong(DocX doc)
        {

            var tableMobileNetwork = doc.AddTable(1, 5);
            tableMobileNetwork.Design = TableDesign.TableGrid;
            tableMobileNetwork.Rows[0].Cells[4].Width = 250;

            tableMobileNetwork.Rows[0].Cells[0].Paragraphs[0].Append("Nhà mạng");
            tableMobileNetwork.Rows[0].Cells[1].Paragraphs[0].Append("Viettel");
            tableMobileNetwork.Rows[0].Cells[2].Paragraphs[0].Append("Mobi");
            tableMobileNetwork.Rows[0].Cells[3].Paragraphs[0].Append("Vina");
            tableMobileNetwork.Rows[0].Cells[4].Paragraphs[0].Append("Tổng ba nhà mạng");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[1].Cells[0].Paragraphs[0].Append("Tổng vị trí");
            tableMobileNetwork.Rows[1].Cells[1].Paragraphs[0].Append("1667");
            tableMobileNetwork.Rows[1].Cells[2].Paragraphs[0].Append("3342");
            tableMobileNetwork.Rows[1].Cells[3].Paragraphs[0].Append("345");
            tableMobileNetwork.Rows[1].Cells[4].Paragraphs[0].Append("6657");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[2].Cells[0].Paragraphs[0].Append("%");
            tableMobileNetwork.Rows[2].Cells[1].Paragraphs[0].Append("19%");
            tableMobileNetwork.Rows[2].Cells[2].Paragraphs[0].Append("26%");
            tableMobileNetwork.Rows[2].Cells[3].Paragraphs[0].Append("55%");
            tableMobileNetwork.Rows[2].Cells[4].Paragraphs[0].Append("100%");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[3].Cells[0].Paragraphs[0].Append("Trạm 2G");
            tableMobileNetwork.Rows[3].Cells[1].Paragraphs[0].Append("845");
            tableMobileNetwork.Rows[3].Cells[2].Paragraphs[0].Append("644");
            tableMobileNetwork.Rows[3].Cells[3].Paragraphs[0].Append("785");
            tableMobileNetwork.Rows[3].Cells[4].Paragraphs[0].Append("2274");
            tableMobileNetwork.InsertRow();

            tableMobileNetwork.Rows[4].Cells[0].Paragraphs[0].Append("%");
            tableMobileNetwork.Rows[4].Cells[1].Paragraphs[0].Append("19%");
            tableMobileNetwork.Rows[4].Cells[2].Paragraphs[0].Append("26%");
            tableMobileNetwork.Rows[4].Cells[3].Paragraphs[0].Append("55%");
            tableMobileNetwork.Rows[4].Cells[4].Paragraphs[0].Append("100%");
            tableMobileNetwork.InsertRow();

            tableMobileNetwork.Rows[5].Cells[0].Paragraphs[0].Append("Trạm 3G");
            tableMobileNetwork.Rows[5].Cells[1].Paragraphs[0].Append("845");
            tableMobileNetwork.Rows[5].Cells[2].Paragraphs[0].Append("644");
            tableMobileNetwork.Rows[5].Cells[3].Paragraphs[0].Append("785");
            tableMobileNetwork.Rows[5].Cells[4].Paragraphs[0].Append("2274");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[6].Cells[0].Paragraphs[0].Append("%");
            tableMobileNetwork.Rows[6].Cells[1].Paragraphs[0].Append("19%");
            tableMobileNetwork.Rows[6].Cells[2].Paragraphs[0].Append("26%");
            tableMobileNetwork.Rows[6].Cells[3].Paragraphs[0].Append("55%");
            tableMobileNetwork.Rows[6].Cells[4].Paragraphs[0].Append("100%");
            tableMobileNetwork.InsertRow();

            tableMobileNetwork.Rows[7].Cells[0].Paragraphs[0].Append("Trạm 4G");
            tableMobileNetwork.Rows[7].Cells[1].Paragraphs[0].Append("845");
            tableMobileNetwork.Rows[7].Cells[2].Paragraphs[0].Append("644");
            tableMobileNetwork.Rows[7].Cells[3].Paragraphs[0].Append("785");
            tableMobileNetwork.Rows[7].Cells[4].Paragraphs[0].Append("2274");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[8].Cells[0].Paragraphs[0].Append("%");
            tableMobileNetwork.Rows[8].Cells[1].Paragraphs[0].Append("19%");
            tableMobileNetwork.Rows[8].Cells[2].Paragraphs[0].Append("26%");
            tableMobileNetwork.Rows[8].Cells[3].Paragraphs[0].Append("55%");
            tableMobileNetwork.Rows[8].Cells[4].Paragraphs[0].Append("100%");
            tableMobileNetwork.InsertRow();

            tableMobileNetwork.Rows[9].Cells[0].Paragraphs[0].Append("Trạm 5G");
            tableMobileNetwork.Rows[9].Cells[1].Paragraphs[0].Append("845");
            tableMobileNetwork.Rows[9].Cells[2].Paragraphs[0].Append("644");
            tableMobileNetwork.Rows[9].Cells[3].Paragraphs[0].Append("785");
            tableMobileNetwork.Rows[9].Cells[4].Paragraphs[0].Append("2274");
            tableMobileNetwork.InsertRow();
            tableMobileNetwork.Rows[10].Cells[0].Paragraphs[0].Append("%");
            tableMobileNetwork.Rows[10].Cells[1].Paragraphs[0].Append("19%");
            tableMobileNetwork.Rows[10].Cells[2].Paragraphs[0].Append("26%");
            tableMobileNetwork.Rows[10].Cells[3].Paragraphs[0].Append("55%");
            tableMobileNetwork.Rows[10].Cells[4].Paragraphs[0].Append("100%");

            for (int row = 0; row < 11; row++)
            {
                for (int cell = 0; cell < 5; cell++)
                {
                    if (row == 0)
                    {
                        tableMobileNetwork.Rows[row].Cells[cell].Paragraphs[0].Bold();
                    }
                    tableMobileNetwork.Rows[row].Cells[cell].Paragraphs[0].Alignment = Alignment.center;
                    if (row % 2 != 0)
                    {
                        tableMobileNetwork.Rows[row].Cells[cell].FillColor = System.Drawing.Color.CornflowerBlue;
                    }
                }
            }
            var p = doc.Paragraphs.Where(s => s.Text.Contains("{banghatangdidong}")).ToList();
            foreach (var item in p)
            {
                item.InsertTableAfterSelf(tableMobileNetwork);
                item.ReplaceText("{banghatangdidong}", "");
            }

        }

        private void bangtruyendan(DocX doc)
        {
            var table = doc.AddTable(2, 9);


            table.Rows[0].MergeCells(1, 4);
            table.Rows[0].MergeCells(2, 5);

            table.MergeCellsInColumn(0, 0, 1);


            for (int i = 0; i < 3; i++)
            {
                table.Rows[0].Cells[i].FillColor = System.Drawing.Color.Yellow;

            }
            for (int i = 1; i < 9; i++)
            {
                table.Rows[1].Cells[i].FillColor = System.Drawing.Color.Yellow;
                table.Rows[1].Cells[i].Width = 200;
            }
            table.Rows[0].Cells[0].Paragraphs[0].Append("Tỉnh").Bold();
            table.Rows[0].Cells[1].Paragraphs[0].Append("Trạm").Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Cáp quang").Bold();

            for (int cell = 0; cell < 3; cell++)
            {
                table.Rows[0].Cells[cell].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[0].Cells[cell].Paragraphs[0].Bold();
            }

            table.Rows[1].Cells[1].Paragraphs[0].Append("Tổng trạm Macro");
            table.Rows[1].Cells[2].Paragraphs[0].Append("Truyền dẫn quang");
            table.Rows[1].Cells[3].Paragraphs[0].Append("Truyền dẫn Viba/Vsat");
            table.Rows[1].Cells[4].Paragraphs[0].Append("Tỷ lệ trạm sử dụng viba,vsat");
            table.Rows[1].Cells[5].Paragraphs[0].Append("Cáp treo (km)");
            table.Rows[1].Cells[6].Paragraphs[0].Append("Cáp ngầm (km)");
            table.Rows[1].Cells[7].Paragraphs[0].Append("Cáp OPGW");
            table.Rows[1].Cells[8].Paragraphs[0].Append("Tổng khối lượng cáp quang (km)");
            for (int cell = 1; cell < 9; cell++)
            {
                table.Rows[1].Cells[cell].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[cell].Paragraphs[0].Bold();
            }

            table.InsertRow();
            table.InsertRow();

            var p = doc.Paragraphs.Where(s => s.Text.Contains("{hatangtruyendan}")).FirstOrDefault();
            if (p != null)
            {
                p.InsertTableAfterSelf(table);
                p.ReplaceText("{hatangtruyendan}", "");
            }

        }

        private void chatluongmangvotuyen(DocX doc, string wordKehoach)
        {
            var table = doc.AddTable(2, 8);
            table.AutoFit = AutoFit.Window;

            table.Rows[0].MergeCells(1, 7);
            table.MergeCellsInColumn(0, 0, 1);

            table.Rows[0].Cells[0].Paragraphs[0].Append("2025-04").FontSize(11).Bold();
            table.Rows[0].Cells[1].Paragraphs[0].Append("Đắk Lắk").FontSize(11).Bold();

            table.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
            table.Rows[0].Cells[1].Paragraphs[0].Alignment = Alignment.center;

            table.Rows[0].Cells[0].FillColor = System.Drawing.Color.Yellow;
            table.Rows[0].Cells[1].FillColor = System.Drawing.Color.Yellow;

            table.Rows[1].Cells[1].Paragraphs[0].Append("CTKT").FontSize(11);
            table.Rows[1].Cells[2].Paragraphs[0].Append("Giá trị đạt được").FontSize(11);
            table.Rows[1].Cells[3].Paragraphs[0].Append("So với CTKT").FontSize(11);
            table.Rows[1].Cells[4].Paragraphs[0].Append("So với tháng trước").FontSize(11);
            table.Rows[1].Cells[5].Paragraphs[0].Append("So với cùng kỳ năm trước").FontSize(11);
            table.Rows[1].Cells[6].Paragraphs[0].Append("Tháng trước").FontSize(11);
            table.Rows[1].Cells[7].Paragraphs[0].Append("Cùng kỳ năm trước").FontSize(11);

            for (int i = 1; i < 8; i++)
            {
                table.Rows[1].Cells[i].Paragraphs[0].Alignment = Alignment.center;
                table.Rows[1].Cells[i].Paragraphs[0].Bold();
                table.Rows[1].Cells[i].FillColor = System.Drawing.Color.Yellow;
            }

            List<string> dataTable = new List<string>();
            string[] columnTable = new string[] { "mvt-A", "mvt-B", "mvt-C", "mvt-D", "mvt-E", "mvt-F", "mvt-G", "mvt-H" };

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordKehoach, false))
            {
                for (int i = 1; i < 17; i++)
                {
                    string rowData = "";
                    foreach (string columns in columnTable)
                    {
                        var sdt = wordDoc.MainDocumentPart.Document.Descendants<SdtElement>()
                      .FirstOrDefault(s =>
                          s.SdtProperties.GetFirstChild<Tag>()?.Val == $"{columns}{i}");
                        if (sdt != null)
                        {
                            rowData += sdt.InnerText + "|";
                        }
                    }

                    dataTable.Add(rowData);
                }

                table.InsertRow();
                int numRow = 2;
                foreach (string data in dataTable)
                {
                    string[] dataSplit = data.Split('|');
                    for (int i = 0; i < dataSplit.Length - 1; i++)
                    {
                        table.Rows[numRow].Cells[i].Paragraphs[0].Append(dataSplit[i]).FontSize(11);
                    }
                    
                    if (numRow < dataTable.Count() - 1)
                    {
                        numRow++;
                        table.InsertRow();

                    }
                }

            }

            var p = doc.Paragraphs.Where(s => s.Text.Contains("{bangchatluongmangvotuyen}")).FirstOrDefault();
            if (p != null)
            {
                p.InsertTableAfterSelf(table);
                p.ReplaceText("{bangchatluongmangvotuyen}", "");
            }
        }

        private void vunglom(DocX doc)
        {
            var table = doc.AddTable(1, 6);

            table.Rows[0].Cells[0].Paragraphs[0].Append("TT").Bold();
            table.Rows[0].Cells[1].Paragraphs[0].Append("Tên tỉnh").Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Tên huyện").Bold();
            table.Rows[0].Cells[3].Paragraphs[0].Append("2G").Bold();
            table.Rows[0].Cells[4].Paragraphs[0].Append("4G").Bold();
            table.Rows[0].Cells[5].Paragraphs[0].Append("Tổng").Bold();
            table.InsertRow();

            var p = doc.Paragraphs.Where(s => s.Text.Contains("{vunglom}")).FirstOrDefault();
            if (p != null)
            {
                p.InsertTableAfterSelf(table);
                p.ReplaceText("{vunglom}", "");
            }

        }

        private void bangLuuluongChatluongmang(DocX doc, string wordKehoach)
        {

            List<string> dataTable = new List<string>();
            string[] rowTable = new string[] { "N", "A", "B", "C", "D", "E", "H", "G", "I", "J", "K", "L" };
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordKehoach, false))
            {

                for (int i = 1; i < 6; i++)
                {
                    string rowData = "";
                    foreach (string row in rowTable)
                    {
                        var sdt = wordDoc.MainDocumentPart.Document.Descendants<SdtElement>()
                      .FirstOrDefault(s =>
                          s.SdtProperties.GetFirstChild<Tag>()?.Val == $"{row}{i}");
                        if (sdt != null)
                        {
                            rowData += sdt.InnerText + "|";
                        }
                    }

                    dataTable.Add(rowData);
                }

            }

            var table = doc.AddTable(2, 12);

            table.AutoFit = AutoFit.Window;
            table.Rows[0].Cells[0].Paragraphs[0].Append("DLK").Bold().FontSize(8);
            table.Rows[0].Cells[1].Paragraphs[0].Append("Tổng lưu lượng thoại/ngày 2G (Erl)").Bold().FontSize(8);
            table.Rows[0].Cells[2].Paragraphs[0].Append("TU HR80%").Bold().FontSize(8);
            table.Rows[0].Cells[3].Paragraphs[0].Append("Tổng lưu lượng thoại/ngày 3G (Erl)").Bold().FontSize(8);
            table.Rows[0].Cells[4].Paragraphs[0].Append("Tổng lưu lượng data/ngày 3G (GB)").Bold().FontSize(8);
            table.Rows[0].Cells[5].Paragraphs[0].Append("Tốc độ 3G (Mbps)").Bold().FontSize(8);
            table.Rows[0].Cells[6].Paragraphs[0].Append("TU 3G Peak (%)").Bold().FontSize(8);
            table.Rows[0].Cells[7].Paragraphs[0].Append("Tổng lưu lượng thoại/ngày 4G (Erl)").Bold().FontSize(8);
            table.Rows[0].Cells[8].Paragraphs[0].Append("Tổng lưu lượng data/ngày 4G (GB)").Bold().FontSize(8);
            table.Rows[0].Cells[9].Paragraphs[0].Append("Tốc độ 4G (Mbps)").Bold().FontSize(8);
            table.Rows[0].Cells[10].Paragraphs[0].Append("TU 4G").Bold().FontSize(8);
            table.Rows[0].Cells[11].Paragraphs[0].Append("Tổng lưu lượng data/ngày 5G (GB)").Bold().FontSize(8);

            for (int i = 0; i < 12; i++)
            {
                table.Rows[0].Cells[i].FillColor = System.Drawing.Color.Yellow;
            }

            table.Rows[1].Cells[1].Paragraphs[0].Append("2G").Bold().FontSize(8);
            table.Rows[1].Cells[3].Paragraphs[0].Append("3G").Bold().FontSize(8);
            table.Rows[1].Cells[7].Paragraphs[0].Append("4G").Bold().FontSize(8);
            table.Rows[1].Cells[11].Paragraphs[0].Append("5G").Bold().FontSize(8);

            table.Rows[1].Cells[1].FillColor = System.Drawing.Color.Yellow;
            table.Rows[1].Cells[3].FillColor = System.Drawing.Color.Yellow;
            table.Rows[1].Cells[7].FillColor = System.Drawing.Color.Yellow;
            table.Rows[1].Cells[11].FillColor = System.Drawing.Color.Yellow;

            table.Rows[1].MergeCells(1, 2);
            table.Rows[1].MergeCells(2, 5);
            table.Rows[1].MergeCells(3, 6);
            table.MergeCellsInColumn(0, 0, 1);

            table.InsertRow();
            int numRow = 2;

            foreach (string data in dataTable)
            {
                string[] dataSplit = data.Split('|');
                for (int i = 0; i < dataSplit.Length; i++)
                {
                    if (i < 12)
                    {
                        table.Rows[numRow].Cells[i].Paragraphs[0].Append(dataSplit[i]).FontSize(8);
                    }
                }
                numRow++;
                table.InsertRow();
            }


            var p = doc.Paragraphs.Where(s => s.Text.Contains("{bangluuluongmangluoi}")).FirstOrDefault();
            if (p != null)
            {
                p.InsertTableAfterSelf(table);
                p.ReplaceText("{bangluuluongmangluoi}", "");
            }
        }

        private void bangchitieuhatang(DocX doc, IXLWorksheet ws, IXLWorksheet wsMeta)
        {

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
            DateTime current = DateTime.Now;
            table.Rows[0].Cells[5].Paragraphs[0].Append($"Thực hiện năm {current.Year}").Font(font).Bold();
            table.Rows[0].Cells[6].Paragraphs[0].Append($"Thực hiện năm {current.Year - 1}").Font(font).Bold();

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

            var t = doc.Paragraphs.Where(s => s.Text.Contains("{bangchitieu}")).FirstOrDefault();
            t.ReplaceText("{bangchitieu}", "");
            t.InsertTableAfterSelf(table);
        }

        private void trienkhaiBTS(DocX doc)
        {
            var tableBTS = doc.AddTable(5, 9);
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

            tableBTS.Rows[0].Cells[4].FillColor = System.Drawing.Color.CornflowerBlue;
            tableBTS.Rows[1].Cells[6].FillColor = System.Drawing.Color.CornflowerBlue;
            tableBTS.Rows[1].Cells[7].FillColor = System.Drawing.Color.CornflowerBlue;
            tableBTS.Rows[1].Cells[8].FillColor = System.Drawing.Color.CornflowerBlue;

            tableBTS.Rows[2].Cells[0].Paragraphs[0].Append("1");
            tableBTS.Rows[3].Cells[0].Paragraphs[0].Append("2");

            tableBTS.Rows[2].Cells[1].Paragraphs[0].Append("DLK");
            tableBTS.Rows[3].Cells[1].Paragraphs[0].Append("PYN");

            tableBTS.Rows[2].Cells[2].Paragraphs[0].Append("20");
            tableBTS.Rows[3].Cells[2].Paragraphs[0].Append("88");

            tableBTS.Rows[4].Cells[1].Paragraphs[0].Append("Tổng");
            var b = doc.Paragraphs.Where(s => s.Text.Contains("{t_trienkhaixaydungbtsmoi}")).First();
            b.ReplaceText("{t_trienkhaixaydungbtsmoi}", "");
            b.InsertTableAfterSelf(tableBTS);
        }

        private void soluongtramtheothuphu(DocX doc)
        {
            // Danh gia so luong tram theo thu phu va nong thon
            var tableDG = doc.AddTable(4, 5);
            tableDG.Design = TableDesign.TableGrid;

            tableDG.Rows[0].Cells[0].Paragraphs[0].Append("Nhà mạng");
            tableDG.Rows[0].Cells[1].Paragraphs[0].Append("Viettel");
            tableDG.Rows[0].Cells[2].Paragraphs[0].Append("Mobi");
            tableDG.Rows[0].Cells[3].Paragraphs[0].Append("Vina");
            tableDG.Rows[0].Cells[4].Paragraphs[0].Append("Tổng ba nhà mạng");

            tableDG.Rows[1].Cells[0].Paragraphs[0].Append("Tổng vị trí");
            tableDG.Rows[1].Cells[1].Paragraphs[0].Append("1667");
            tableDG.Rows[1].Cells[2].Paragraphs[0].Append("3342");
            tableDG.Rows[1].Cells[3].Paragraphs[0].Append("345");
            tableDG.Rows[1].Cells[4].Paragraphs[0].Append("6657");

            tableDG.Rows[2].Cells[0].Paragraphs[0].Append("Thủ phủ");
            tableDG.Rows[2].Cells[1].Paragraphs[0].Append("6665");
            tableDG.Rows[2].Cells[2].Paragraphs[0].Append("998");
            tableDG.Rows[2].Cells[3].Paragraphs[0].Append("557");
            tableDG.Rows[2].Cells[4].Paragraphs[0].Append("54");

            tableDG.Rows[3].Cells[0].Paragraphs[0].Append("Nông thôn");
            tableDG.Rows[3].Cells[1].Paragraphs[0].Append("845");
            tableDG.Rows[3].Cells[2].Paragraphs[0].Append("644");
            tableDG.Rows[3].Cells[3].Paragraphs[0].Append("785");
            tableDG.Rows[3].Cells[4].Paragraphs[0].Append("2274");

            var m = doc.Paragraphs.FirstOrDefault(s => s.Text.Contains("{soluongtramtheothuphu}"));
            m.InsertTableAfterSelf(tableDG);
            m.ReplaceText("{soluongtramtheothuphu}", "");
        }
        public void generateWordFile(IProgress<string> progress, string fileTheoDoi = "", string sheetTheoDoi = "", string wordkehoach = "")
        {
            progress.Report("Đang tạo báo cáo word");

            var wb = new XLWorkbook(fileTheoDoi);
            var ws = wb.Worksheet("BC_chi_tiet");

            var wsMeta = wb.Worksheet("Meta");
            List<string> mainCreatia = new List<string>();

            string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Baocao.docx");

            using (var doc = DocX.Load(wordPath))
            {

                //doc.ReplaceText("{thang}", "08");
                doc.ReplaceText("{nam}", DateTime.Now.Year.ToString());

                /*doc.ReplaceText("{nhanxet01}", "Vị trí trạm hiện tại Viettel đang chiếm ưu thế với 1633 vị trí. Số lượng vị trí trạm Viettel nhiều hơn Vinaphone 240 vị trí và nhiều hơn Mobifone 510 vị trí. Xét về mức huyện Viettel còn 4 huyện có vị trí trạm ít hơn nhà mạng Vina là Krông Bông ít hơn 4 vị trí, Huyện Ea Súp và Krông Búk ít hơn 1 vị trí, huyện Ea Súp ít hơn 5 trạm");
                doc.ReplaceText("{h_ketquathuchien6thang}", "KẾT QUẢ THỰC HIỆN 6 THÁNG ĐẦU NĂM 2025");*/

                /*hatangdidong(doc);
                bangchitieuhatang(doc, ws, wsMeta);
                soluongtramtheothuphu(doc);
                trienkhaiBTS(doc);
                bangtruyendan(doc);
                vunglom(doc);*/
                //  bangLuuluongChatluongmang(doc, wordkehoach);
                chatluongmangvotuyen(doc, wordkehoach);
                wb.Dispose();

                doc.Save();
                doc.Dispose();
            }
            progress.Report("Tạo file word hoàn tất");
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
