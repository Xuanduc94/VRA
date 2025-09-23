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

            /* if(workbook.Worksheets.FirstOrDefault(c => c.Name =="MetaTH")== null)
             {
                 workbook.Worksheets.Add("MetaTH");
             }
 */
            var workSheet = workbook.Worksheet("BC_chi_tiet");

            //var wsTh = workbook.Worksheet("MetaTH");


            var workSheetMeta = workbook.Worksheet(7);
            int count = workSheet.RowsUsed().Count();
            int rowMeta = 3;
            int id = 1;
            for (int row = 3; row <= count; row++)
            {
                workSheetMeta.Cell("A" + rowMeta).Value = id;
                string Id = StringHelper.RemoveDiacriticsAndSpaces(workSheet.Cell("B" + row).Value.ToString().ToLower());
                workSheetMeta.Cell("B" + rowMeta).Value = Id;
                workSheet.Cell("Q" + row).Value = Id;

                rowMeta++;
                id++;
            }
            workbook.Save();
            workbook.Dispose();
            progress.Report("Ánh xạ thành công");
        }

    }
}
