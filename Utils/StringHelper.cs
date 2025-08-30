using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Utils
{
    public static class StringHelper
    {
       
        public static string RemoveDiacriticsAndSpaces(string text, bool removeSpace = true)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;
            string normalized = text.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            foreach (char c in normalized)
            {
                UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(c);
                if (removeSpace)
                {
                    if (uc != UnicodeCategory.NonSpacingMark && c != ' ')
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    if (uc != UnicodeCategory.NonSpacingMark)
                    {
                        sb.Append(c);
                    }
                }
                

            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}
