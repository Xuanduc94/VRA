using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Controllers
{
    public class PowerPointController
    {
        public void generateSildeFile(IProgress<string> progress, string file)
        {
            progress.Report("Tạo silde trình chiếu");
        }
    }
}
