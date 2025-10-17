using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNRHistoryNext
{
    public class History
    {
        public Remarks PNRRemarks;
        public History(string pnr)
        {
            ReadPNRHistoryList(pnr);
        }
        void ReadPNRHistoryList(string pnr)
        {
            // TODO
            PNRRemarks = new Remarks();
            //SqlConnection cnn = new System.Data.SqlClient.SqlConnection(NextDB.ATPIConnections.TWS);
            //cnn.Open();
            //SqlCommand cmd = cnn.CreateCommand();
            //cmd.Parameters.Add(@"@PNR", System.Data.SqlDbType.NVarChar, 6).Value = pnr;
            //cmd.CommandType = System.Data.CommandType.Text;
            //cmd.CommandText = $@"SELECT [CrsPNRID]
            //                    ,[CrsHIST]
            //                    FROM [TWS].[dbo].[Bookings_CRS]
            //                    WHERE CrsPNRID IN ({pnr})";
            //SqlDataReader reader = cmd.ExecuteReader();
            //PNRRemarks = new Remarks();
            //while (reader.Read())
            //{
            //    PNRRemarks.AddRemarks(reader["CrsPNRID"].ToString(), reader["CrsHIST"].ToString());
            //}
            //reader.Close();
            //cnn.Close();

        }
    }
}
