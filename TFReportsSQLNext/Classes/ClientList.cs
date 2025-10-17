using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQLNext.Classes
{
    public class ClientList : List<Client>
    {
        public void Load()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new(AtpiGrDBConnections.PanasoftConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd ;
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
    public class ClientDataSet
    {
        public Microsoft.Data.SqlClient.SqlCommand ClientDataSetCmd()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new(AtpiGrDBConnections.PanasoftConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd;
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientList";
            return sqlCmd;
        }
    }
    public class Client
    {
        public required string ClientCode { get; set; }
        public required string ClientName { get; set; }
        public required string DispName { get; set; } 
        public override string ToString()
        {
            return $"{DispName}";
        }
    }
}
