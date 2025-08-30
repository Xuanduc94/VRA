using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Viettel_Report_Automation.Config;
using Viettel_Report_Automation.Models;
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
                //processWordExcel(fileWord, fileExcel, progress);

            }
            catch (Exception ex)
            {

                throw;
            }

        }

        private void TonghopTheoDoi(string fileChamdiem, string fileTheodoi, string fileWord, IProgress<string> progress)
        {
            progress.Report("Trích xuất file theo dõi KPI");
            var workbookChamDiem = new XLWorkbook(fileChamdiem);
            var sheetChamDiem = workbookChamDiem.Worksheet("BC_chi_tiet");
            string monthStr = StringHelper.RemoveDiacriticsAndSpaces(sheetChamDiem.Cell("G1").Value.ToString(), false).Replace("/", ".").ToUpper();
            progress.Report("Cấu hình theo dõi");

            /* progress.Report("Cấu hình hoàn tất");
             this.CreateMeTracking(fileTheodoi, monthStr.Trim());
             progress.Report("Tính toán dữ liệu");
             WriteToKPIFile(fileTheodoi, fileChamdiem);*/
            string m = sheetChamDiem.Cell("G1").Value.ToString();
            string[] strSpilt = m.Split('/', ' ');
            int.TryParse(strSpilt[1], out int month);
            QuarterlyReport(progress, fileTheodoi, month);
            workbookChamDiem.Dispose();

        }

        private void CreateMeTracking(string fileTheodoi, string month)
        {

            var workbook = new XLWorkbook(fileTheodoi);
            var worksheet = workbook.Worksheet(month.Trim());
            if (workbook.Worksheets.FirstOrDefault(s => s.Name == "Meta") != null)
            {
                workbook.Worksheets.Delete("Meta");
            }
            workbook.AddWorksheet("Meta");
            var sheetMeta = workbook.Worksheet("Meta");
            int rowMeta = 1;
            for (int row = 5; row < worksheet.RowsUsed().Count(); row++)
            {
                var data = worksheet.Cell("B" + row).Value.ToString();
                if (data != "")
                {
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

        public void QuarterlyReport(IProgress<string> progress, string fileTheodoi, int month)
        {

            progress.Report("Đang tạo báo cáo quý");

            try
            {
                var workbookTong = new XLWorkbook(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx"));
                int quater = NumberHelper.GetQuarter(month);
                int[] monthOfQuater = new int[3];
                switch (quater)
                {
                    case 1:
                        monthOfQuater[0] = 1;
                        monthOfQuater[1] = 2;
                        monthOfQuater[2] = 3;
                        break;
                    case 2:
                        monthOfQuater[0] = 4;
                        monthOfQuater[1] = 5;
                        monthOfQuater[2] = 6;
                        break;
                    case 3:
                        monthOfQuater[0] = 7;
                        monthOfQuater[1] = 8;
                        monthOfQuater[2] = 9;
                        break;
                    case 4:
                        monthOfQuater[0] = 10;
                        monthOfQuater[1] = 11;
                        monthOfQuater[2] = 12;
                        break;
                    default:
                        break;
                }
                var workbookTheoDoi = new XLWorkbook(fileTheodoi);

                bool check = true;
                foreach (int i in monthOfQuater)
                {
                    var ws = workbookTheoDoi.Worksheets.FirstOrDefault(s => s.Name == "THANG " + i + "." + DateTime.Now.Year);
                    if (ws == null)
                    {
                        check = false;
                        break;
                    }
                }
                if (check == false)
                {
                    progress.Report("Không đủ dữ liệu tạo báo cáo quý");
                    return;
                }
                // Tong hop bao cao

                foreach (int i in monthOfQuater)
                {
                    var ws = workbookTheoDoi.Worksheets.FirstOrDefault(s => s.Name == "THANG " + i + "." + DateTime.Now.Year);

                }
            }
            catch (Exception)
            {

                throw;
            }

        }

    }
}
