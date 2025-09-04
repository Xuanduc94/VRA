using ClosedXML.Excel;
using System.Windows;
using Viettel_Report_Automation.Utils;

namespace Viettel_Report_Automation.Controllers
{
    public class SettingController
    {
        // Tao Id cho file cham diem
        public void SettingScore(string _fileExcel, IProgress<string> progress)
        {
            progress.Report("Đang ánh xạ chỉ số");
            var workbook = new XLWorkbook(_fileExcel);

            if(workbook.Worksheets.FirstOrDefault(c => c.Name =="MetaTH")== null)
            {
                workbook.Worksheets.Add("MetaTH");
            }

            var workSheet = workbook.Worksheet("BC_chi_tiet");

            var wsTh = workbook.Worksheet("MetaTH");


            var workSheetMeta = workbook.Worksheet(7);
            int count = workSheet.RowsUsed().Count();
            int rowMeta  = 3;
            int id = 1;
            for (int row = 3; row <= count; row++)
            {
                workSheetMeta.Cell("A" + rowMeta).Value = id;
                string Id = StringHelper.RemoveDiacriticsAndSpaces(workSheet.Cell("B" + row).Value.ToString());
                workSheetMeta.Cell("B" + rowMeta).Value =Id;

                workSheet.Cell(row, 7).FormulaA1 = ("=Meta!C"+ rowMeta);
                workSheet.Cell(row, 8).FormulaA1 = "=Meta!D" + rowMeta;

                workSheet.Cell("J" + row).FormulaA1 = "=Meta!E" + rowMeta;
                workSheet.Cell("K" + row).FormulaA1 = "=Meta!F" + rowMeta;

                workSheet.Cell("M" + row).FormulaA1 = "=Meta!G" + rowMeta;
                workSheet.Cell("N" + row).FormulaA1 = "=Meta!H" + rowMeta;

                wsTh.Cell($"A{rowMeta}").Value = Id;
                workSheet.Cell($"D{row}").FormulaA1 = $"=MetaTH!B{rowMeta}";
                workSheet.Cell($"E{row}").FormulaA1 = $"=MetaTH!C{rowMeta}";
                workSheet.Cell($"F{row}").FormulaA1 = $"=MetaTH!D{rowMeta}";
                rowMeta++;  
                id++;
            }
            workbook.Save();
            workbook.Dispose();
            progress.Report("Ánh xạ thành công");
        }
        
    }
}
