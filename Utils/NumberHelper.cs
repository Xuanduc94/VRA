using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Utils
{
    public static class NumberHelper
    {
        public static double ParseStringToDouble(string _number)
        {
            if (double.TryParse(_number, out double result))
            {
               
                if (double.IsNaN(result) || double.IsInfinity(result))
                    return 0;

                return Math.Round(result, 2);
            }

            return 0;
        }

        public static int GetQuarter(int month)
        {
            if (month < 1 || month > 12)
                throw new ArgumentOutOfRangeException("month", "Tháng phải nằm trong khoảng 1-12");

            return (month - 1) / 3 + 1;
        }
    }
}
