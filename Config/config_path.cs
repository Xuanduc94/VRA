using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Config
{
    public static class config_path
    {
        public static string _chamdiemPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files","ChamdiemKPI.xlsx") ;

        public static string _tonghopquyexcel = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files", "Tonghopquy.xlsx");
    }
}
