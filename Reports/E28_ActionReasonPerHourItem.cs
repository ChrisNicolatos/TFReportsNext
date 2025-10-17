using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public class E28_ActionReasonPerHourItem
    {
        public int Index { get; set; }
        public int Hour { get; set; }
        public int DayPart { get; set; } // 0=day 1=Evening 2=Night 3=Morning
        public int Count { get; set; }
        public int DelayTotal { get; set; }
        public decimal DelayAverage()
        {
            if (Count == 0) return 0;
            return (decimal)DelayTotal / (decimal)Count;
        }
    }
}
