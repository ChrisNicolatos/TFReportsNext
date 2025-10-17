using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQL.Classes
{
    public class ClientList:List<Client>
    {
        public void Load()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new Microsoft.Data.SqlClient.SqlConnection(NextDB.DBConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd = new Microsoft.Data.SqlClient.SqlCommand();
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientList";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                this.Add(new Client
                {
                    ClientCode = sqlRdr.GetString(0),
                    ClientName = sqlRdr.GetString(1),
                    DispName = sqlRdr.GetString(2)
                });
            }
        }
    }
    public class Client
    {
        public string ClientCode { get; set; }
        public string ClientName { get; set; }
        public string DispName { get; set; }
        public override string ToString()
        {
            return $"{ClientCode} - {DispName}";
        }
    }
}
