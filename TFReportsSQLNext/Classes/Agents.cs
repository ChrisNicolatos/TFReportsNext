using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQLNext.Classes
{
    public class Agents : List<string>
    {
        public void Load()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new(AtpiGrDBConnections.PanasoftConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd ;
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCmd.CommandText = "ATPIData.dbo.TFReports_AgentGroups";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                Add(sqlRdr.GetString(0));
            }
        }
    }
}
