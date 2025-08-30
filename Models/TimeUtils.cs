using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Models
{
    public static class TimeUtils
    {
        public static int GetQuarter(int month)
        {
            return (month - 1) / 3 + 1;
        }

        public static bool IsQuarterComplete(IEnumerable<int> months, int quarter)
        {
            var quarterMonths = new Dictionary<int, int[]>
        {
            {1, new[]{1, 2, 3}},
            {2, new[]{4, 5, 6}},
            {3, new[]{7, 8, 9}},
            {4, new[]{10, 11, 12}},
        };

            return quarterMonths[quarter].All(m => months.Contains(m));
        }
    }
}
