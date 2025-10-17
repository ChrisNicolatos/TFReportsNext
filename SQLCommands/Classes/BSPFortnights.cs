using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQL
{
    public class BSPFortnights : List<DateTime>
    {
        public void Load()
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new Microsoft.Data.SqlClient.SqlConnection(NextDB.DBConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd = new Microsoft.Data.SqlClient.SqlCommand();
            sqlCmd=sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCmd.CommandText = "ATPIData.dbo.TFReports_BSPFortnights";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                this.Add(  sqlRdr.GetDateTime(0));
            }
        }
    }
    
}
