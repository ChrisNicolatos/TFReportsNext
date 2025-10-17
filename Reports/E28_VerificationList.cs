using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public class E28_VerificationList : Dictionary<string, E28_VerificationItem>
    {
        public E28_VerificationItem Totals = new E28_VerificationItem("Grand Total");
        public void Add(string verification, int delaytotal, decimal actualsavings, decimal potentialsavings, string farecurrency, string downsellcurrency)
        {
            if (!this.ContainsKey(verification))
            {
                E28_VerificationItem item = new E28_VerificationItem(verification);
                this.Add(verification, item);
            }
            this[verification].Count += 1;
            this[verification].DelayTotal += delaytotal;
            Totals.Count += 1;
            Totals.DelayTotal += delaytotal;
            if (farecurrency == downsellcurrency)
            {
                this[verification].ActualSavings += actualsavings;
                this[verification].PotentialSavings += potentialsavings;
                Totals.ActualSavings += actualsavings;
                Totals.PotentialSavings += potentialsavings;
            }

        }
    }
}
