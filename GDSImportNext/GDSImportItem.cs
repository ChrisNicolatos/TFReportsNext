using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GDSImportNext
{
    public class GDSImportItem
    {
        private int gdsdataid;
        public int Id { get; set; }
        public int GDSDataID
        {
            get
            {
                return gdsdataid;
            }
            set
            {
                gdsdataid = value;
                ClientReferences = new GDSImportClientReferences();
                ClientReferences.Read(gdsdataid);
            }
        }
        public DateTime ImportTimeStamp { get; set; }
        public DateTime AutoErrorTimeStamp { get; set; }
        public string BookOfficeId { get; set; }
        public string PNR { get; set; }
        public string BookSalesman { get; set; }
        public string IssueSalesman { get; set; }
        public string Product { get; set; }
        public string ImportStatus { get; set; }
        public string ImportMessage { get; set; }
        public string ImportType { get; set; }
        public string TransactionType { get; set; }
        public string Type { get; set; }
        public string PNRStatus { get; set; }
        public string PassengerSNs { get; set; }
        public string ClientCode { get; set; }
        public string ClientName { get; set; }
        public string VesselRequired { get; set; }
        public string Vessel { get; set; }
        public bool IsPassive { get; set; }
        public string ChargeType { get; set; }
        public string Routing { get; set; }
        public string Supplier { get; set; }
        public string Number { get; set; }
        public string Passengers { get; set; }
        public DateTime DepartureDate { get; set; }
        public string OwnedFileName { get; set; }
        public GDSImportClientReferences ClientReferences { get; private set; }
        public bool HasErrors
        {
            get
            {
                if (TransactionType != "Refund" && (ImportStatus == "Import Failed" || ImportMessage != "" || ImportType.EndsWith("*") || ClientReferences.Errors.Count > 0))
                {
                    return true;

                }
                else { return false; }

            }
        }

    }
}
