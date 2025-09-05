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
                
                doc.ReplaceText("{nhanxet01}", "Vị trí trạm hiện tại Viettel đang chiếm ưu thế với 1633 vị trí. Số lượng vị trí trạm Viettel nhiều hơn Vinaphone 240 vị trí và nhiều hơn Mobifone 510 vị trí. Xét về mức huyện Viettel còn 4 huyện có vị trí trạm ít hơn nhà mạng Vina là Krông Bông ít hơn 4 vị trí, Huyện Ea Súp và Krông Búk ít hơn 1 vị trí, huyện Ea Súp ít hơn 5 trạm");
                doc.Save();
            }
                progress.Report("Đã tạo file word");
        }
    }
}
