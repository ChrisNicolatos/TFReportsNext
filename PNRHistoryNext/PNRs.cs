using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRHistoryNext
{
    public class PNRs : List<PNR>
    {
        public SortedDictionary<string, PNRTotals> TotsPerBookAgent = new SortedDictionary<string, PNRTotals>();
        public SortedDictionary<string, PNRTotals> TotsPerIssueAgent = new SortedDictionary<string, PNRTotals>();
        public SortedDictionary<string, PNRTotals> TotsPerClient = new SortedDictionary<string, PNRTotals>();
        public SortedDictionary<string, PNRTotals> TotsPerAgentGroup = new SortedDictionary<string, PNRTotals>();
        public PNRTotals GrandTotals = new PNRTotals();
        List<PNR> templist;

        public void ReadPNRs(DateTime datefrom, DateTime dateto)
        {
            SqlConnection cnn = new Microsoft.Data.SqlClient.SqlConnection(AtpiGrDBConnections.PanasoftConnections.PNR);
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            cmd.Parameters.Add("@TransDateFrom", SqlDbType.Date).Value = datefrom;
            cmd.Parameters.Add("@TransDateTo", SqlDbType.Date).Value = dateto;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 3000;
            cmd.CommandText = "ATPIData.dbo.TFReports_E32_QCSellingPrice";

            SqlDataReader reader = cmd.ExecuteReader();

            string allpnrs = "";
            this.Clear();
            templist = new List<PNR>();
            TotsPerAgentGroup = new SortedDictionary<string, PNRTotals>();
            TotsPerBookAgent = new SortedDictionary<string, PNRTotals>();
            TotsPerClient = new SortedDictionary<string, PNRTotals>();
            TotsPerIssueAgent = new SortedDictionary<string, PNRTotals>();
            GrandTotals = new PNRTotals();

            while (reader.Read())
            {
                PNR pnr = new PNR()
                {
                    GDS = reader["GDS"].ToString(),
                    PNRId = reader["PNR"].ToString(),
                    PNRCreationDate = DateTime.MinValue,
                    TransactionDate = DateTime.MinValue,
                    ClientCode = reader["ClientCode"].ToString(),
                    ClientName = reader["ClientName"].ToString(),
                    AgentGroup = reader["AgentGroup"].ToString(),
                    BookPCC = reader["BookOfficeId"].ToString(),
                    BookSalesman = reader["BookSalesman"].ToString(),
                    IssueOID = reader["IssueOfficeId"].ToString(),
                    IssueSalesman = reader["Issuing Salesman"].ToString(),
                    SellingPrice = reader["Selling Price"].ToString().Trim(),
                    DummyRemarks = reader["Visa-Dummy-Protect"].ToString().Trim(),
                    FixedMarkupClient = Convert.ToBoolean(reader["FixedMarkUpClient"]),
                    Void = Convert.ToBoolean(reader["Void"]),
                    Refund = Convert.ToBoolean(reader["Refund"]),
                    Passive = Convert.ToBoolean(reader["IsPassive"]),
                    ProductTypeId = Convert.ToInt32(reader["TypeID"]),
                    ProductName = reader["ProductType"].ToString()
                };
                if (DateTime.TryParse(reader["PNRCreationDate"].ToString(), out DateTime creat))
                {
                    pnr.PNRCreationDate = creat;
                }
                if (DateTime.TryParse(reader["TransactionDate"].ToString(), out DateTime trans))
                {
                    pnr.TransactionDate = trans;
                }
                templist.Add(pnr);
                if (allpnrs.Length > 0) { allpnrs += ","; }
                allpnrs += $"'{pnr.PNRId}'";
            }
            reader.Close();
            cnn.Close();

            ReadHistory(allpnrs);
            CalculateTotals();
        }
        void ReadHistory(string allpnrs)
        {
            History history = new History(allpnrs);
            foreach (PNR item in templist)
            {
                string hist = "";
                string histdummy = "";
                string clientcode = item.ClientCode;
                if (history.PNRRemarks.ContainsKey(item.PNRId))
                {
                    hist += history.PNRRemarks[item.PNRId];//.FreeFlow;

                }
                if (history.PNRRemarks.DummyRemarks.ContainsKey(item.PNRId))
                {
                    histdummy += history.PNRRemarks.DummyRemarks[item.PNRId];
                }
                if (clientcode == "" && history.PNRRemarks.ClientCode.ContainsKey(item.PNRId))
                {
                    clientcode = history.PNRRemarks.ClientCode[item.PNRId];
                }
                this.Add(new PNR
                {
                    GDS = item.GDS,
                    PNRId = item.PNRId,
                    PNRCreationDate = item.PNRCreationDate,
                    TransactionDate = item.TransactionDate,
                    ClientCode = clientcode, // item.ClientCode,
                    ClientName = item.ClientName,
                    AgentGroup = item.AgentGroup,
                    BookPCC = item.BookPCC,
                    BookSalesman = item.BookSalesman,
                    IssueOID = item.IssueOID,
                    IssueSalesman = item.IssueSalesman,
                    SellingPrice = item.SellingPrice,
                    DummyRemarks = item.DummyRemarks,
                    FromHistory = hist,
                    FromHistoryDummy = histdummy,
                    FixedMarkupClient = item.FixedMarkupClient,
                    Void = item.Void,
                    Refund = item.Refund,
                    Passive = item.Passive,
                    ProductTypeId = item.ProductTypeId,
                    ProductName = item.ProductName
                });
            }
        }
        void CalculateTotals()
        {
            bool WithoutSellingPrice;
            foreach (PNR item in this)
            {
                WithoutSellingPrice = true;
                PNRTotals totsissue;
                PNRTotals totsbook;
                PNRTotals totsclient;
                PNRTotals totsagent;
                if (!TotsPerIssueAgent.ContainsKey(item.IssueSalesman)) { totsissue = new PNRTotals(); TotsPerIssueAgent.Add(item.IssueSalesman, totsissue); }
                if (!TotsPerBookAgent.ContainsKey(item.BookSalesman)) { totsbook = new PNRTotals(); TotsPerBookAgent.Add(item.BookSalesman, totsbook); }
                if (!TotsPerClient.ContainsKey(item.Client)) { totsclient = new PNRTotals(); TotsPerClient.Add(item.Client, totsclient); }
                if (!TotsPerAgentGroup.ContainsKey(item.AgentGroup)) { totsagent = new PNRTotals(); TotsPerAgentGroup.Add(item.AgentGroup, totsagent); }
                totsissue = TotsPerIssueAgent[item.IssueSalesman];
                totsbook = TotsPerBookAgent[item.BookSalesman];
                totsclient = TotsPerClient[item.Client];
                totsagent = TotsPerAgentGroup[item.AgentGroup];

                if (item.FixedMarkupClient)
                {
                    WithoutSellingPrice = false;
                    totsissue.FixedMarkupClient++;
                    totsbook.FixedMarkupClient++;
                    totsclient.FixedMarkupClient++;
                    totsagent.FixedMarkupClient++;
                    GrandTotals.FixedMarkupClient++;
                }
                if (item.Void)
                {
                    WithoutSellingPrice = false;
                    totsissue.Void++;
                    totsbook.Void++;
                    totsclient.Void++;
                    totsagent.Void++;
                    GrandTotals.Void++;
                }
                if (item.Refund)
                {
                    WithoutSellingPrice = false;
                    totsissue.Refund++;
                    totsbook.Refund++;
                    totsclient.Refund++;
                    totsagent.Refund++;
                    GrandTotals.Refund++;
                }
                if (item.Passive)
                {
                    WithoutSellingPrice = false;
                    totsissue.Passive++;
                    totsbook.Passive++;
                    totsclient.Passive++;
                    totsagent.Passive++;
                    GrandTotals.Passive++;
                }
                if (!item.IsAirTicket)
                {
                    WithoutSellingPrice = false;
                    totsissue.NotAirTicket++;
                    totsbook.NotAirTicket++;
                    totsclient.NotAirTicket++;
                    totsagent.NotAirTicket++;
                    GrandTotals.NotAirTicket++;
                }
                if (item.AllDummyRemarks.Trim().Length > 0)
                {
                    WithoutSellingPrice = false;
                    totsissue.DummyRemarks++;
                    totsbook.DummyRemarks++;
                    totsclient.DummyRemarks++;
                    totsagent.DummyRemarks++;
                    GrandTotals.DummyRemarks++;
                }
                if (item.AllSPRemarks.Trim().Length > 0)
                {
                    WithoutSellingPrice = false;
                    totsissue.WithSell++;
                    totsbook.WithSell++;
                    totsclient.WithSell++;
                    totsagent.WithSell++;
                    GrandTotals.WithSell++;
                }
                if (WithoutSellingPrice)
                {
                    totsissue.WithoutSell++;
                    totsbook.WithoutSell++;
                    totsclient.WithoutSell++;
                    totsagent.WithoutSell++;
                    GrandTotals.WithoutSell++;
                }
                TotsPerIssueAgent[item.IssueSalesman] = totsissue;
                TotsPerBookAgent[item.BookSalesman] = totsbook;
                TotsPerClient[item.Client] = totsclient;
                TotsPerAgentGroup[item.AgentGroup] = totsagent;
            }
        }
    }
    public class PNRTotals
    {
        public int DummyRemarks { get; set; }
        public int FixedMarkupClient { get; set; }
        public int Void { get; set; }
        public int Refund { get; set; }
        public int Passive { get; set; }
        public int NotAirTicket { get; set; }
        public int WithSell { get; set; }
        public int WithoutSell { get; set; }
        public int Total
        {
            get
            {
                return DummyRemarks + FixedMarkupClient + Void + Refund + Passive + WithSell + WithoutSell;
            }
        }
        public int TotalSaleable
        {
            get
            {
                return WithSell + WithoutSell;
            }
        }
        public decimal WithSellPct
        {
            get
            {
                if (TotalSaleable == 0)
                {
                    return 0;
                }
                else
                {
                    return (decimal)WithSell / (decimal)TotalSaleable;
                }
            }
        }
        public decimal WithoutSellPct
        {
            get
            {
                if (TotalSaleable == 0)
                {
                    return 0;
                }
                else
                {
                    return (decimal)WithoutSell / (decimal)TotalSaleable;
                }
            }
        }
    }
}
