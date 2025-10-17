using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public class E28_VerificationItem
    {
        public E28_VerificationItem(string verification)
        {
            Verification = verification;
            Count = 0;
            DelayTotal = 0;
            ActualSavings = 0;
            PotentialSavings = 0;
        }
        public string Verification { get; set; }
        public int Count { get; set; }
        public int DelayTotal { get; set; }
        public decimal ActualSavings { get; set; }
        public decimal PotentialSavings { get; set; }
        public decimal AverageDelay
        {
            get
            {
                if (Count == 0) return 0;
                return (decimal)DelayTotal / (decimal)Count;
            }
        }
    }
}
