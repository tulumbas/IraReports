using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IraReports.Models.Source
{
    class AdRecord
    {
        private string _durationString;

        public DateTime Date { get; set; }
        public string Time { get; set; }
        public string Code { get; set; }
        public string DurationString
        {
            get
            {
                return _durationString;
            }
            set
            {
                _durationString = value;
                TimeSpan ts;
                if(TimeSpan.TryParse(value, out ts))
                {
                    Duration = ts;
                }
            }
        }

        public TimeSpan Duration { get; private set; }
        public AdInfo Info { get; set; }        
    }
}
