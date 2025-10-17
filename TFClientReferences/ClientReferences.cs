using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFClientReferences
{
    public class ClientReferences:Dictionary<string, ClientReference>
    {
        public ClientReferences(string accountcode) { 

        }
        void Read(string accountcode)
        {
            using (SqlConnection conn = new SqlConnection(NextDB.DBConnections.TravelForce))
            {
                SqlCommand comm = new SqlCommand();
                int sequencenumber = 0;
                ClientReferenceGDSEntryCollection gdsentries = new ClientReferenceGDSEntryCollection();

                conn.Open();
                comm = conn.CreateCommand();
                comm.CommandType = System.Data.CommandType.Text;
                comm.Parameters.Add("@EntityID", System.Data.SqlDbType.Int).Value = entityId;
                comm.CommandText = @" SELECT ClientCustomProperties.Id  
,ClientCustomProperties.CustomPropertyID
,ClientCustomProperties.LookUpValues
,ClientCustomProperties.LimitToLookUp
,ClientCustomProperties.Label
,ClientCustomProperties.TFEntityID
,ClientCustomProperties.LTRequiredTypeID
,CustomProperties.LTAssociatedWithID
FROM TFC_ClientCustomProperties ClientCustomProperties
LEFT JOIN TFC_CustomProperties CustomProperties
ON ClientCustomProperties.CustomPropertyID = CustomProperties.Id
WHERE ClientCustomProperties.TFEntityID = @EntityID AND CustomProperties.LTKindID = 691
AND ClientCustomProperties.IsDisabled = 0";
                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    //add reference field for Travel Force
                    sequencenumber++;
                    if (gdsentries.TryGetValue("199", out ClientReferenceGDSEntry gdsrefentry))
                    {
                        ClientReference item = new ClientReference(sequencenumber, "99", "Reference", gdsrefentry.Amadeus, gdsrefentry.Galileo, false, false, "", false, CLIENTREFINPUTMODE.PerBooking);
                        this.Add(item.SequenceNo, item);
                        ValueIndex.Add(item.Id, item);
                    }
                    while (reader.Read())
                    {
                        sequencenumber++;
                        bool mandatory = false;
                        CLIENTREFINPUTMODE inputmode = CLIENTREFINPUTMODE.PerBooking;
                        if ((int)reader["LTRequiredTypeID"] == (int)NextUtilities.EnumItems.ClientReferenceRequiredType.PropertyReqToSave || (int)reader["LTRequiredTypeID"] == (int)NextUtilities.EnumItems.ClientReferenceRequiredType.PropertyReqToInv)
                        {
                            mandatory = true;
                        }
                        if ((int)reader["LTAssociatedWithID"] == (int)NextUtilities.EnumItems.ClientReferenceConnects.Passenger)
                        {
                            inputmode = CLIENTREFINPUTMODE.PerPax;
                        }
                        if (gdsentries.TryGetValue((100 + (int)reader["CustomPropertyID"]).ToString(), out gdsrefentry))
                        {
                            ClientReference item = new ClientReferencesNext.ClientReference(sequencenumber, reader["CustomPropertyID"].ToString(), reader["Label"].ToString(), gdsrefentry.Amadeus, gdsrefentry.Galileo, mandatory, false, reader["LookUpValues"].ToString(), Convert.ToBoolean(reader["LimitToLookUp"]), inputmode);
                            this.Add(item.SequenceNo, item);
                            ValueIndex.Add(item.Id, item);
                        }
                    }
                    reader.Close();
                }
                conn.Close();
            }
        }
    }
}
