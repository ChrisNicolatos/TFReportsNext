using Reports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQL.Classes
{
    public class ClientGroupList:List<ClientGroup>
    {
        public void Load( Reports.ReportsItem.ClientCodeSelect GroupsType)
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new Microsoft.Data.SqlClient.SqlConnection(NextDB.DBConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd = new Microsoft.Data.SqlClient.SqlCommand();
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            if (GroupsType == Reports.ReportsItem.ClientCodeSelect.ClientCodeAndSeachefsGroup)
                sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientGroups";
            else
                sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientGroupsAll";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                this.Add(new ClientGroup
                {
                    Id = sqlRdr.GetInt32(0),
                    GroupName = sqlRdr.GetString(1)
                });
            }
        }
    }
    
    public class ClientGroup
    {
        public int Id { get; set; } = 0;
        public string GroupName { get; set; } = "";
        public override string ToString()
        {
            return GroupName;
        }
    }
}
