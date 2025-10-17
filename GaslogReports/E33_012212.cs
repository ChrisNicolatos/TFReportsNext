using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace GaslogReports
{
    public class E33_Item
    {
        public int Id { get; set; }
        public string ClientCode { get; set; }
        public string ClientName { get; set; }
        public DateTime IssueDate { get; set; }
        public string TransactionType { get; set; }
        public string ActionType { get; set; }
        public string ReasonForTravel { get; set; }
        public string RFTHeading
        {
            get
            {
                string temp = ReasonForTravel;
                switch (ReasonForTravel)
                {
                    case "ED-ONSIGNER":
                    case "ED-OFFSIGNER":
                        temp = "ED-ONSIGNER/OFFSIGNER";
                        break;
                    default:
                        break;
                }
                return temp;
            }
        }
        public string Reference { get; set; }
        public string Vessel { get; set; }
        public string ProductType { get; set; }
        public string PNR { get; set; }
        public string TicketNumber { get; set; }
        public string Passenger { get; set; }
        public string InvCode { get; set; }
        public string InvSeries { get; set; }
        public long InvNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string BookedBy { get; set; }
        public decimal NetPayable { get; set; }
        public string Verified { get; set; }
        public string Remarks { get; set; }
        public int RegNr { get; set; }
        public string TicketingAirline { get; set; }
        public string Routing { get; set; }
        public string SalesPerson { get; set; }
        public string IssuingAgent { get; set; }
        public string CreatorAgent { get; set; }
        public DateTime DepartureDate { get; set; }
        public DateTime ArrivalDate { get; set; }
        public string ConnectedDocument { get; set; }
        public int DocStatusID { get; set; }
        public string CancelsDoc { get; set; }
        public string ServicesDescription { get; set; }
        public string SortKey
        {
            get
            {
                return $"{TransactionKey}{RegNr}{Id}";
            }
        }
        public string TransactionKey
        {
            get
            {
                string temp = "";
                if (ActionType == "Issue") { temp += "0"; } else { temp += "1"; }
                switch (ReasonForTravel)
                {
                    case "CA-TRAINING COMP":
                        temp += $"01{Reference}";
                        break;
                    case "ED-ONSIGNER":
                    case "ED-OFFSIGNER":
                        temp += $"02{Vessel}";
                        break;
                    case "CA-HSSE":
                        temp += $"03{Vessel}";
                        break;
                    case "UKRANIAN CRISIS":
                        temp += $"04{Vessel}";
                        break;
                    case "CHAIRMAN AWARDS":
                        temp += $"05{Vessel}";
                        break;
                    case "REC-BRIEFING/RECRUITING":
                        temp += $"06{Vessel}";
                        break;
                    default:
                        temp += "000";
                        break;
                }


                return temp;
            }
        }
    }
    public class E33_Collection : SortedDictionary<string, E33_Item>
    {
        public E33_Collection(DateTime datefrom, DateTime dateto, string ClientCode)
        {
            SqlConnection cnn = new System.Data.SqlClient.SqlConnection(NextDB.PanasoftConnections.PNR);
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            cmd.Parameters.Add("@FromDate", SqlDbType.Date).Value = datefrom;
            cmd.Parameters.Add("@ToDate", SqlDbType.Date).Value = dateto;
            cmd.Parameters.Add("@ClientCode", SqlDbType.NVarChar,20).Value = ClientCode;
            cmd.CommandText = "ATPIData.dbo.TFReports_E33_012212_CrewPreparationForInvoicing";

            SqlDataReader reader = cmd.ExecuteReader();
            int id = 0;
            while (reader.Read())
            {
                id++;
                E33_Item item = new E33_Item()
                {
                    Id = id,
                    ClientCode = reader.GetString(0),
                    ClientName = reader.GetString(1),
                    IssueDate = reader.GetDateTime(2),
                    TransactionType = reader.GetString(3),
                    ActionType = reader.GetString(4),
                    ReasonForTravel = reader.GetString(5),
                    Reference = reader.GetString(6),
                    Vessel = reader.GetString(7),
                    ProductType = reader.GetString(8),
                    PNR = reader.GetString(9),
                    TicketNumber = reader.GetString(10),
                    Passenger = reader.GetString(11),
                    InvCode = reader.GetString(12),
                    InvSeries = reader.GetString(13),
                    InvNumber = reader.GetInt64(14),
                    InvoiceDate = reader.GetDateTime(15),
                    BookedBy = reader.GetString(16),
                    NetPayable = reader.GetDecimal(17),
                    Verified = reader.GetString(18),
                    Remarks = reader.GetString(19),
                    RegNr = reader.GetInt32(20),
                    TicketingAirline = reader.GetString(21),
                    Routing = reader.GetString(22),
                    SalesPerson = reader.GetString(23),
                    IssuingAgent = reader.GetString(24),
                    CreatorAgent = reader.GetString(25),
                    DepartureDate = reader.GetDateTime(26),
                    ArrivalDate = reader.GetDateTime(27),
                    ConnectedDocument = reader.GetString(28),
                    DocStatusID = reader.GetInt32(29),
                    CancelsDoc = reader.GetString(30),
                    ServicesDescription = reader.GetString(31)
                };
                this.Add($"{item.SortKey}", item);
            }
            reader.Close();
            cnn.Close();
        }
    }
}
