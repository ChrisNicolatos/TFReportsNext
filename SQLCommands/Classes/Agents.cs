using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQL
{
    public class Agents:List<string>
    {
        public void Load()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new Microsoft.Data.SqlClient.SqlConnection(NextDB.DBConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd = new Microsoft.Data.SqlClient.SqlCommand();
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCmd.CommandText = "ATPIData.dbo.TFReports_AgentGroups";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                this.Add(sqlRdr.GetString(0));
            }
        }
    }
}
