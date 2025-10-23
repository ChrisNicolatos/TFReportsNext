using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDSImportNext
{
    public class GDSImportItems : List<GDSImportItem>
    {
        public void Read()
        {
            SqlConnection cnn = new Microsoft.Data.SqlClient.SqlConnection(AtpiGrDBConnections.PanasoftConnections.PNR);
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            cmd.CommandText = "ATPIData.dbo.TFReports_E39_GDSImportErrors";

            SqlDataReader reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                GDSImportItem item = new GDSImportItem()
                {
                    Id = Convert.ToInt32(reader["Id"]),
                    GDSDataID = Convert.ToInt32(reader["GDSDataID"]),
                    ImportTimeStamp = DateTime.MinValue,
                    AutoErrorTimeStamp = DateTime.MinValue,
                    BookOfficeId = reader["BookOfficeId"].ToString(),
                    PNR = reader["PNR"].ToString(),
                    BookSalesman = reader["BookSalesman"].ToString(),
                    IssueSalesman = reader["IssueSalesman"].ToString(),
                    Product = reader["Product"].ToString(),
                    ImportStatus = reader["ImportStatus"].ToString(),
                    ImportMessage = reader["ImportMessage"].ToString(),
                    ImportType = reader["ImportType"].ToString(),
                    TransactionType = reader["TransactionType"].ToString(),
                    Type = reader["Type"].ToString(),
                    PNRStatus = reader["PNRStatus"].ToString(),
                    PassengerSNs = reader["PassengerSNs"].ToString(),
                    ClientCode = reader["ClientCode"].ToString(),
                    ClientName = reader["ClientName"].ToString(),
                    VesselRequired = reader["VesselRequired"].ToString(),
                    Vessel = reader["Vessel"].ToString(),
                    IsPassive = Convert.ToBoolean(reader["IsPassive"]),
                    ChargeType = reader["ChargeType"].ToString(),
                    Routing = reader["Routing"].ToString(),
                    Supplier = reader["Supplier"].ToString(),
                    Number = reader["Number"].ToString(),
                    Passengers = reader["Passengers"].ToString(),
                    DepartureDate = DateTime.MinValue,
                    OwnedFileName = reader["OwnedFileName"].ToString()
                };
                if (DateTime.TryParse(reader["ImportTimeStamp"].ToString(), out DateTime import))
                {
                    item.ImportTimeStamp = import;
                }
                if (DateTime.TryParse(reader["AutoErrorTimeStamp"].ToString(), out DateTime autoerror))
                {
                    item.AutoErrorTimeStamp = autoerror;
                }
                if (DateTime.TryParse(reader["DepartureDate"].ToString(), out DateTime departure))
                {
                    item.DepartureDate = departure;
                }

                this.Add(item);
            }
            reader.Close();
            cnn.Close();
        }
    }
}
