using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Models
{
    public class MergeRange
    {
        public MergeRange(IXLRangeAddress address, int rowSpan, int colSpan, string value)
        {
            this.address = address;
            this.rowSpan = rowSpan;
            this.colSpan = colSpan;
            this.value = value;
        }

        public IXLRangeAddress address { get; set; }
        public int rowSpan { get; set; }
        public int colSpan { get; set; }
        public string value { get; set; }
    }
}
