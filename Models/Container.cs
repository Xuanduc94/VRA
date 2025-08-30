using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Models
{
    public class Container
    {
        public Container(string wordField, string excelField)
        {
            this.wordField = wordField;
            this.excelField = excelField;
        }

        public string wordField { get; set; }
        public string excelField { get; set; }
    }
}
