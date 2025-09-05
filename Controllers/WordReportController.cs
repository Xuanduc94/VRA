using System.IO;
using Xceed.Words.NET;

namespace Viettel_Report_Automation.Controllers
{
    public class WordReportController
    {

        public void generateWordFile(IProgress<string> progress)
        {
            string wordPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "word_mau.docx");

            using (var doc = DocX.Load(wordPath))
            {
                doc.ReplaceText("{thang}", "08");
                doc.ReplaceText("{nam}", "2025");
                doc.Save();
            }
                progress.Report("Đã tạo file word");
        }
    }
}
