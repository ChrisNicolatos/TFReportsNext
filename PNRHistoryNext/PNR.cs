using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRHistoryNext
{
    public class PNR
    {
        private string sellingprice;
        private string dummyremarks;

        private string fromhistory;
        private string fromhistorydummy;
        private string allspremarks;
        private string alldummyremarks;
        private bool remarksspdirty;
        private bool remarksdummydirty;

        public string GDS { get; set; }
        public string PNRId { get; set; }
        public DateTime PNRCreationDate { get; set; }
        public DateTime TransactionDate { get; set; }
        public string ClientCode { get; set; }
        public string ClientName { get; set; }
        public string Client
        {
            get
            {
                if (ClientCode + ClientName != "")
                {
                    return $"{ClientCode} - {ClientName}";
                }
                else
                {
                    return "";
                }
            }
        }
        public string AgentGroup { get; set; }
        public string BookPCC { get; set; }
        public string BookSalesman { get; set; }
        public string IssueOID { get; set; }
        public string IssueSalesman { get; set; }
        public bool FixedMarkupClient { get; set; }
        public bool Void { get; set; }
        public bool Refund { get; set; }
        public bool Passive { get; set; }
        public int ProductTypeId { get; set; }
        public string ProductName { get; set; }
        public bool IsAirTicket
        {
            get
            {
                return (ProductName.ToUpper() == "TICKET" || ProductTypeId == 323 || ProductTypeId == 838);
            }
        }
        public string SellingPrice
        {
            get
            {
                return sellingprice;
            }
            set
            {
                sellingprice = value.Trim();
                remarksspdirty = true;
            }
        }
        public string DummyRemarks
        {
            get
            {
                return dummyremarks;
            }
            set
            {
                dummyremarks = value.Trim();
                remarksdummydirty = true;
            }
        }
        public string FromHistory
        {
            get
            {
                return fromhistory;
            }
            set
            {
                fromhistory = value.Trim();
                remarksspdirty = true;
            }
        }
        public string FromHistoryDummy
        {
            get
            {
                return fromhistorydummy;
            }
            set
            {
                fromhistorydummy = value.Trim();
                remarksdummydirty = true;
            }
        }
        public string AllSPRemarks
        {
            get
            {
                if (remarksspdirty)
                {
                    FixSPRemarks();
                }
                return allspremarks;
            }
        }
        void FixSPRemarks()
        {
            allspremarks = "";
            string[] temp = (sellingprice + "|" + fromhistory).Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i <= temp.GetUpperBound(0); i++)
            {
                if (allspremarks.IndexOf(temp[i]) < 0)
                {
                    allspremarks += "|" + temp[i].Trim();
                }
            }
            while (allspremarks.StartsWith("|"))
            {
                allspremarks = allspremarks.Substring(1);
            }
            remarksspdirty = false;

        }
        public string AllDummyRemarks
        {
            get
            {
                if (remarksdummydirty)
                {
                    FixDummyRemarks();
                }
                return alldummyremarks;
            }
        }
        void FixDummyRemarks()
        {
            alldummyremarks = "";
            string[] temp = (dummyremarks + "|" + fromhistorydummy).Split("|".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i <= temp.GetUpperBound(0); i++)
            {
                if (alldummyremarks.IndexOf(temp[i]) < 0)
                {
                    alldummyremarks += "|" + temp[i].Trim();
                }
            }
            while (alldummyremarks.StartsWith("|"))
            {
                alldummyremarks = alldummyremarks.Substring(1);
            }
            remarksdummydirty = false;

        }
    }
}
