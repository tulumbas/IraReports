using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IraReports.Models.Core
{
    static class Utils
    {
        public static bool TryParseDate(string s, out DateTime dt)
        {
            if (DateTime.TryParse(s, out dt))
            {
                return true;
            }

            double val;
            if (double.TryParse(s, out val))
            {
                dt = DateTime.FromOADate(val);
                return true;
            }

            return false;
        }

        public static string SanitizeFileName(string text)
        {
            return Regex.Replace(text, @"[\\\/\:\. ;'""]", "_");
        }
    }
}
