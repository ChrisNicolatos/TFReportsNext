using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDSImport
{
    public class GDSImportClientReferences:List<GDSImportClientReference>
    {
        public List<GDSImportClientReference> Errors { get; set; }
        public void Read(int dataid)
        {
            SqlConnection cnn = new System.Data.SqlClient.SqlConnection(NextDB.PanasoftConnections.PNR);
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@GDSDataId", SqlDbType.BigInt).Value = dataid;
            cmd.CommandText = "ATPIData.dbo.TFReports_E39_GDSImportClientReferences";

            SqlDataReader reader = cmd.ExecuteReader();
            this.Clear();
            Errors = new List<GDSImportClientReference>();

            while (reader.Read())
            {
                GDSImportClientReference item = new GDSImportClientReference()
                {
                    Id = Convert.ToInt32(reader["Id"]),
                    Element = reader["Element"].ToString(),
                    Description = reader["Description"].ToString(),
                    Required = reader["Required"].ToString(),
                    LimitToLookUp = Convert.ToBoolean(reader["LimitToLookup"]),
                    RawValue = reader["RawValue"].ToString(),
                    Value = reader["Value"].ToString(),
                    ValueFound = reader["ValueFound"].ToString(),
                    Found = reader["Found"].ToString()
                };
                this.Add(item);
                if(!item.Found.StartsWith("OK")) Errors.Add(item);
            }
            reader.Close();
            cnn.Close();
        }
    }
}
