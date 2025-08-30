using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Models
{
   public class KPI
    {
        public KPI(string name, float value)
        {
            Name = name;
            Value = value;
        }

        public string Name { get; set; }
        public float Value { get; set; }
    }
}
