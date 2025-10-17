using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GaslogReports
{
    public class E35_Item
    {
        public int Id { get; set; }
        public int RegNr { get; set; }
        public string BookedBy { get; set; }
        public string Office { get; set; }
        public string CostCentre { get; set; }
        public string ReasonForTravel { get; set; }
        public decimal NetPayable { get; set; }
        public string Routing { get; set; }
        public string Passenger { get; set; }
        public string ProductType { get; set; }
        public string ActionType { get; set; }
        public DateTime IssueDate { get; set; }
        public string Verified { get; set; }
        public string TransactionType { get; set; }
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
                string temp0 = "";
                string temp1 = "";
                if (ActionType == "Issue") { temp0 = "0"; } else { temp0 = "1"; }
                temp1 = $"{BookedBy}|{Office}|{CostCentre}";
                switch ($"{CostCentre}0000".Substring(0, 4))
                {
                    case "3033":
                    case "3034":
                    case "3035":
                        temp += $"{temp0}1|{temp1}|{ReasonForTravel}";
                        break;
                    default:
                        temp += $"{temp0}0|{temp1}";
                        break;
                }
                return temp;
            }
        }
    }
    public class E35_Collection : SortedDictionary<string, E35_Item>
    {
        public E35_Collection(DateTime datefrom, DateTime dateto, string ClientCode )
        {
            SqlConnection cnn = new System.Data.SqlClient.SqlConnection(NextDB.PanasoftConnections.PNR);
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            cmd.Parameters.Add("@FromDate", SqlDbType.Date).Value = datefrom;
            cmd.Parameters.Add("@ToDate", SqlDbType.Date).Value = dateto;
            cmd.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ClientCode;
            cmd.CommandText = "ATPIData.dbo.TFReports_E35_012217_CrewPreparationForInvoicing";

            SqlDataReader reader = cmd.ExecuteReader();
            int id = 0;
            while (reader.Read())
            {
                id++;
                E35_Item item = new E35_Item()
                {
                    Id = id,
                    RegNr = reader.GetInt32(0),
                    BookedBy = reader.GetString(1),
                    Office = reader.GetString(2),
                    CostCentre = reader.GetString(3),
                    ReasonForTravel = reader.GetString(4),
                    NetPayable = reader.GetDecimal(5),
                    Routing = reader.GetString(6),
                    Passenger = reader.GetString(7),
                    ProductType = reader.GetString(8),
                    ActionType = reader.GetString(9),
                    IssueDate = reader.GetDateTime(10),
                    Verified = reader.GetString(11),
                    TransactionType = reader.GetString(12),
                };
                this.Add($"{item.SortKey}", item);
            }
            reader.Close();
            cnn.Close();
        }
    }
}

