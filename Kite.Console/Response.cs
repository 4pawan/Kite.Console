using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Zerodha.Excel
{
    public class Candles
    {
        public DateTime Date { get; set; }
        public string DateFormated { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public double PrevDayClose { get; set; }
        public long Volume { get; set; }
        public double DayLowToHigh { get; set; }
        public double CentHighFrmY { get; set; } // high - PrevDayClose
        public double CentLowFrmY { get; set; }  // PrevDayClose- low
        public double CentCloseFrmY { get; set; }  // close -PrevDayClose OneDayChange
        public double DayCentLowToHigh { get; set; }
        public double Gap { get; set; } // Open- PrevDayClose
        public string _10AM { get; set; }
        public string _10_30AM { get; set; }
        public string _1PM { get; set; }
        public string _2PM { get; set; }
        public string DayMaxHighReachedAt { get; set; }
        public string DayMaxLowReachedAt { get; set; }
        public string IstWeeklyDay { get; set; }

    }

    public class Response
    {
        public string status { get; set; }
        public Data data { get; set; }
    }
    public class Data
    {
        public List<List<object>> candles { get; set; }
    }

}
