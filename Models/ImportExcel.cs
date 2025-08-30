using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Models
{
    public class ImportExcel
    {
        public ImportExcel(string cell, float value)
        {
            _cell = cell;
            _value = value;
        }

        public string _cell { get; set; }
        public float _value { get; set; }
    }
}
