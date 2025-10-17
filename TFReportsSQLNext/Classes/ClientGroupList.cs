using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQLNext.Classes
{
    public class ClientGroupList : List<ClientGroup>
    {
        public void Load(ReportsNext.ReportsItem.ClientCodeSelect GroupsType)
        {
            Microsoft.Data.SqlClient.SqlConnection sqlConn = new(AtpiGrDBConnections.PanasoftConnections.ATPIData);
            sqlConn.Open();
            Microsoft.Data.SqlClient.SqlCommand sqlCmd ;
            sqlCmd = sqlConn.CreateCommand();
            sqlCmd.CommandType = System.Data.CommandType.StoredProcedure;
            if (GroupsType == ReportsNext.ReportsItem.ClientCodeSelect.ClientCodeAndSeachefsGroup)
                sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientGroupsSeaChefs";
            else
                sqlCmd.CommandText = "ATPIData.dbo.TFReports_ClientGroupsAll";
            Microsoft.Data.SqlClient.SqlDataReader sqlRdr = sqlCmd.ExecuteReader();
            while (sqlRdr.Read())
            {
                this.Add(new ClientGroup
                {
                    Id = sqlRdr.GetInt64(0),
                    GroupName = sqlRdr.GetString(1)
                });
            }
        }
    }

    public class ClientGroup
    {
        public long Id { get; set; } = 0;
        public string GroupName { get; set; } = "";
        public override string ToString()
        {
            return GroupName;
        }
    }
}
