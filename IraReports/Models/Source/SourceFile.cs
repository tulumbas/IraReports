using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IraReports.Models.Source
{
    class SourceFile
    {
        public string FileName { get; set; }
        public List<AdRecord> Records { get; }
        public int Count { get { return Records.Count; } }

        public string Canal { get; internal set; }
        public string Date { get; internal set; }

        public SourceFile()
        {
            Records = new List<AdRecord>();
        }

        public void AddRange(IEnumerable<AdRecord> records)
        {
            Records.AddRange(records);
        }

        public double GetTotalSeconds()
        {
            return Records.Sum(x => x.Duration.TotalSeconds);
        }

        public string CreateName(string client)
        {           
            var name = SanitizeFileName($"{client}__{Canal}") ;
            return name + ".xls";
        }

        private string SanitizeFileName(string text)
        {
            return Regex.Replace(text, @"[\\\/\:\. ;'""]", "_");
        }
    }
}
