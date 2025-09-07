using ClosedXML.Excel;
using System.IO;
using System.Windows.Input;
using Viettel_Report_Automation.Utils;

namespace Viettel_Report_Automation.Controllers
{
    public class ReportExtractController
    {


        public void generateReport(string fileChamdiem, string fileTheodoi, string fileWord, IProgress<string> progress)
        {

            try
            {
                TonghopTheoDoi(fileChamdiem, fileTheodoi, fileWord, progress);

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void TonghopTheoDoi(string fileChamdiem, string fileTheodoi, string fileWord, IProgress<string> progress)
        {
           progress.Report("Trích xuất file theo dõi KPI");
            var workbookChamDiem = new XLWorkbook(fileChamdiem);
            var sheetChamDiem = workbookChamDiem.Worksheet("BC_chi_tiet");
            string monthStr = StringHelper.RemoveDiacriticsAndSpaces(sheetChamDiem.Cell("G1").Value.ToString(), false).Replace("/", ".").ToUpper();
            progress.Report("Tiến hành cấu hình");
            this.CreateMeTracking(fileTheodoi, monthStr.Trim());
            progress.Report("Tính toán dữ liệu");
            WriteToKPIFile(fileTheodoi, fileChamdiem);
            string m = sheetChamDiem.Cell("G1").Value.ToString();
            string[] strSpilt = m.Split('/', ' ');
            int.TryParse(strSpilt[1], out int month);
            progress.Report("Tổng hợp báo cáo quý");
            QuarterlyReport(progress, fileTheodoi, month);
            workbookChamDiem.Dispose();
            progress.Report("Tính toán báo cáo năm");
            YearReport();
            progress.Report("Tổng hợp báo cáo");
            int quater = NumberHelper.GetQuarter(month);
            MappingDataYearAndQuaterToKPI(quater, fileChamdiem);
            MappingDataTotal(fileChamdiem);
            progress.Report("Đã tính toán xong");
        }

        private void MappingDataTotal(string fileChamDiem)
        {
            var workbookChamDiem = new XLWorkbook(fileChamDiem);
            var sheetChamDiem = workbookChamDiem.Worksheet("MetaTH");
            var wbTong = new XLWorkbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx"));
            var yearSheet = wbTong.Worksheet("Nam");
            for (int row = 3; row < yearSheet.RowsUsed().Count(); row++)
            {
                string keyword = yearSheet.Cell($"A{row}").Value.ToString();
                var cell = sheetChamDiem.Cells().FirstOrDefault(c => c.GetString() == keyword);
                if (cell != null)
                {
                    int rowCell = cell.Address.RowNumber;
                    // Du lieu nam 
                    sheetChamDiem.Cell($"B{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"C{row}").Value.ToString());
                    sheetChamDiem.Cell($"C{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"E{row}").Value.ToString());
                    sheetChamDiem.Cell($"D{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"G{row}").Value.ToString());

                }
            }

            workbookChamDiem.Save();
            workbookChamDiem.Dispose();
            wbTong.Dispose();
        }

        private void MappingDataYearAndQuaterToKPI(int quater, string fileChamDiem)
        {
            var workbookChamDiem = new XLWorkbook(fileChamDiem);
            var sheetChamDiem = workbookChamDiem.Worksheet("Meta");
            var wbTong = new XLWorkbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx"));
            var yearSheet = wbTong.Worksheet("Nam");
            // Mapping quater data
            for (int row = 3; row < yearSheet.RowsUsed().Count(); row++)
            {
                string keyword = yearSheet.Cell($"A{row}").Value.ToString();
                var cell = sheetChamDiem.Cells().FirstOrDefault(c => c.GetString() == keyword);
                if (cell != null)
                {
                    int rowCell = cell.Address.RowNumber;

                    // Du lieu nam 
                    sheetChamDiem.Cell($"G{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"J{row}").Value.ToString());
                    sheetChamDiem.Cell($"H{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"K{row}").Value.ToString());

                    switch (quater)
                    {
                        case 1:
                            sheetChamDiem.Cell($"E{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"B{row}").Value.ToString());
                            sheetChamDiem.Cell($"F{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"C{row}").Value.ToString());

                            break;
                        case 2:
                            sheetChamDiem.Cell($"E{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"D{row}").Value.ToString());
                            sheetChamDiem.Cell($"F{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"E{row}").Value.ToString());

                            break;
                        case 3:
                            sheetChamDiem.Cell($"E{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"F{row}").Value.ToString());
                            sheetChamDiem.Cell($"F{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"G{row}").Value.ToString());

                            break;
                        case 4:
                            sheetChamDiem.Cell($"E{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"H{row}").Value.ToString());
                            sheetChamDiem.Cell($"F{rowCell}").Value = NumberHelper.ParseStringToDouble(yearSheet.Cell($"I{row}").Value.ToString());
                            break;
                    }
                }

            }
            workbookChamDiem.Save();
            workbookChamDiem.Dispose();
            wbTong.Dispose();

        }

        private void CreateMeTracking(string fileTheodoi, string month)
        {

            var workbook = new XLWorkbook(fileTheodoi);
            var worksheet = workbook.Worksheet(month.Trim());
            string m = month.Replace("THANG ", "");
            if (workbook.Worksheets.FirstOrDefault(s => s.Name == "Meta" + m) != null)
            {
                workbook.Worksheets.Delete("Meta" + m);
            }

            workbook.AddWorksheet("Meta" + m);
            var sheetMeta = workbook.Worksheet("Meta" + m);
            int rowMeta = 1;
            for (int row = 5; row < worksheet.RowsUsed().Count(); row++)
            {
                var data = worksheet.Cell("B" + row).Value.ToString();
                if (data != "")
                {
                    // Bỏ hết dấu và dấu cách để làm Id
                    sheetMeta.Cell("A" + rowMeta).Value = StringHelper.RemoveDiacriticsAndSpaces(data);
                    sheetMeta.Cell("B" + rowMeta).Value = NumberHelper.ParseStringToDouble(worksheet.Cell("G" + row).Value.ToString());
                    sheetMeta.Cell("C" + rowMeta).Value = NumberHelper.ParseStringToDouble(worksheet.Cell("H" + row).Value.ToString());
                    rowMeta++;
                }
            }
            workbook.Save();
            workbook.Dispose();
        }

        private void WriteToKPIFile(string fileTheodoi, string fileChamdiem)
        {
            var workbook = new XLWorkbook(fileTheodoi);
            if (workbook.Worksheets.FirstOrDefault(s => s.Name == "Meta") == null)
            {
                workbook.Worksheets.Add("Meta");
            }

            var worksheet = workbook.Worksheet("Meta");
            var workbookKPI = new XLWorkbook(fileChamdiem);
            var worksheetKPI = workbookKPI.Worksheet("Meta");

            for (int row = 1; row < worksheet.RowsUsed().Count(); row++)
            {
                string findStr = worksheet.Cell("A" + row).Value.ToString();
                int rowKPI = findValue(worksheetKPI, findStr);
                if (rowKPI != 0)
                {
                    worksheetKPI.Cell("C" + rowKPI).Value = NumberHelper.ParseStringToDouble(worksheet.Cell("B" + row).Value.ToString());
                    worksheetKPI.Cell("D" + rowKPI).Value = NumberHelper.ParseStringToDouble(worksheet.Cell("C" + row).Value.ToString());
                }
            }
            workbookKPI.Save();
            workbook.Dispose();
            workbookKPI.Dispose();
        }

        private int findValue(IXLWorksheet workSheet, string keyword)
        {
            int result = 0;
            for (int row = 1; row < workSheet.RowsUsed().Count(); row++)
            {
                string key = workSheet.Cell("B" + row).Value.ToString();
                if (key == keyword)
                {
                    result = row; break;
                }
            }
            return result;
        }

        private void QuarterlyReport(IProgress<string> progress, string fileTheodoi, int month)
        {

            progress.Report("Đang tạo báo cáo quý");

            try
            {
                var workbookTong = new XLWorkbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx"));
                int quater = NumberHelper.GetQuarter(month);
                IXLWorksheet worksheetTong = null;
                int[] monthOfQuater = new int[3];
                switch (quater)
                {
                    case 1:
                        monthOfQuater[0] = 1;
                        monthOfQuater[1] = 2;
                        monthOfQuater[2] = 3;
                        worksheetTong = workbookTong.Worksheet("Quy I");
                        break;
                    case 2:
                        monthOfQuater[0] = 4;
                        monthOfQuater[1] = 5;
                        monthOfQuater[2] = 6;
                        worksheetTong = workbookTong.Worksheet("Quy II");
                        break;
                    case 3:
                        monthOfQuater[0] = 7;
                        monthOfQuater[1] = 8;
                        monthOfQuater[2] = 9;
                        worksheetTong = workbookTong.Worksheet("Quy III");
                        break;
                    case 4:
                        monthOfQuater[0] = 10;
                        monthOfQuater[1] = 11;
                        monthOfQuater[2] = 12;
                        worksheetTong = workbookTong.Worksheet("Quy IV");
                        break;
                    default:
                        break;
                }
                var workbookTheoDoi = new XLWorkbook(fileTheodoi);

                bool check = true;
                List<int> months = new List<int>();
                foreach (int i in monthOfQuater)
                {
                    var ws = workbookTheoDoi.Worksheets.FirstOrDefault(s => s.Name == "THANG " + i + "." + DateTime.Now.Year);
                    if (ws != null)
                    {
                        months.Add(i);
                    }
                }

                // Tong hop bao cao
                // Lay bao cao theo thang 
                foreach (int i in months)
                {
                    int step = 1;
                    var ws = workbookTheoDoi.Worksheets.FirstOrDefault(s => s.Name == "Meta" + i + "." + DateTime.Now.Year);

                    // Ghi du lieu cac thang
                    if (ws != null)
                    {
                        int rowTh = 3;
                        for (int row = 1; row < ws.RowsUsed().Count(); row++)
                        {
                            string keyword = ws.Cell("A" + row).Value.ToString();
                            //var search = worksheetTong.Search(keyword);
                            var cell = worksheetTong.Cells().FirstOrDefault(c => c.GetString() == keyword);
                            // Ghi đầu mục chỉ tiêu
                            if (cell == null)
                            {
                                worksheetTong.Cell($"A{rowTh}").Value = keyword;
                                worksheetTong.Cell($"H{rowTh}").FormulaA1 = $"=B{rowTh}+D{rowTh}+F{rowTh}";
                                worksheetTong.Cell($"I{rowTh}").FormulaA1 = $"=C{rowTh}+E{rowTh}+G{rowTh}";
                                rowTh++;
                            }
                        }

                        for (int row = 1; row < ws.RowsUsed().Count(); row++)
                        {
                            string keyword = ws.Cell("A" + row).Value.ToString();
                            var cell = worksheetTong.Cells().FirstOrDefault(c => c.GetString() == keyword);
                            var rowNumber = cell.Address.RowNumber;
                            switch (step)
                            {
                                case 1:
                                    worksheetTong.Cell("B" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("B" + row).Value.ToString());
                                    worksheetTong.Cell("C" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("C" + row).Value.ToString());
                                    break;
                                case 2:
                                    worksheetTong.Cell("D" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("B" + row).Value.ToString());
                                    worksheetTong.Cell("E" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("C" + row).Value.ToString());
                                    break;
                                case 3:
                                    worksheetTong.Cell("F" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("B" + row).Value.ToString());
                                    worksheetTong.Cell("G" + rowNumber).Value = NumberHelper.ParseStringToDouble(ws.Cell("C" + row).Value.ToString());
                                    break;
                            }

                        }
                    }
                    step++;
                }
                workbookTong.Save();
                workbookTong.Dispose();
                workbookTheoDoi.Dispose();
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        private void YearReport()
        {
            var wbTong = new XLWorkbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx"));
            var yearSheet = wbTong.Worksheet("Nam");
            // Ghi cac chi muc vo sheet tong hop nam
            List<string> quaterSheet = new List<string>();
            quaterSheet.Add("Quy I");
            quaterSheet.Add("Quy II");
            quaterSheet.Add("Quy III");
            quaterSheet.Add("Quy IV");
            int rowYearSheet = 3;
            foreach (string quater in quaterSheet)
            {
                var worksheetQuater = wbTong.Worksheet(quater);
                for (int rowQuater = 3; rowQuater < worksheetQuater.RowsUsed().Count(); rowQuater++)
                {
                    var checkExist = yearSheet.Cells().FirstOrDefault(c => c.GetString() == worksheetQuater.Cell($"A{rowQuater}").Value.ToString());
                    if (checkExist == null)
                    {
                        yearSheet.Cell($"A{rowYearSheet}").Value = worksheetQuater.Cell($"A{rowQuater}").Value.ToString();
                        rowYearSheet++;
                    }
                }
            }

            // Ghi dữ liệu vô sheet tổng hợp năm
            foreach (string quater in quaterSheet)
            {
                var worksheetQuater = wbTong.Worksheet(quater);
                for (int rowQuater = 3; rowQuater < worksheetQuater.RowsUsed().Count(); rowQuater++)
                {
                    string index = worksheetQuater.Cell("A" + rowQuater).Value.ToString();
                    var cell = yearSheet.Cells().FirstOrDefault(c => c.GetString() == index);
                    int rowCell = cell.Address.RowNumber;
                    switch (quater)
                    {
                        case "Quy I":
                            yearSheet.Cell($"B{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"H{rowQuater}").Value.ToString());
                            yearSheet.Cell($"C{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"I{rowQuater}").Value.ToString());
                            break;
                        case "Quy II":
                            yearSheet.Cell($"D{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"H{rowQuater}").Value.ToString());
                            yearSheet.Cell($"E{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"I{rowQuater}").Value.ToString());
                            break;
                        case "Quy III":
                            yearSheet.Cell($"F{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"H{rowQuater}").Value.ToString());
                            yearSheet.Cell($"G{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"I{rowQuater}").Value.ToString());
                            break;
                        case "Quy IV":
                            yearSheet.Cell($"H{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"H{rowQuater}").Value.ToString());
                            yearSheet.Cell($"I{rowCell}").Value = NumberHelper.ParseStringToDouble(worksheetQuater.Cell($"I{rowQuater}").Value.ToString());
                            break;

                    }
                    yearSheet.Cell($"J{rowCell}").FormulaA1 = $"=B{rowQuater}+D{rowQuater}+F{rowQuater}+H{rowQuater}";
                    yearSheet.Cell($"K{rowCell}").FormulaA1 = $"=C{rowQuater}+E{rowQuater}+G{rowQuater}+I{rowQuater}";
                }
            }
            wbTong.Save();
            wbTong.Dispose();

        }
    }
}
