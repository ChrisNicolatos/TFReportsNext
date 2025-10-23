using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFSpreadSheetsNext
{
    public partial class TFRCommon
    {
        public string E57_TUI_030366()
        {
            int RowCounter;
            int LastColumn = 15;

            Totals? Statement = null;
            Totals? Client = null;
            Totals? Vessel = null;
            Totals? Staff = null;

            bool WithStaffTotal = mobjReports.BooleanOption1;
            Statement = new Totals("Statement", $"INVOICE SUMMARY {mobjReports.Date1To:yyyyMMdd}", "");

            xlWorkSheet.FreezePanes(1, 0);
            xlWorkSheet.SetColumnStyle(2, mStyles.xlStyleDate);
            xlWorkSheet.SetColumnStyle(5, mStyles.xlStyleInteger);
            xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDate);
            xlWorkSheet.SetColumnStyle(12, 15, mStyles.xlStyleDecimal);
            RowCounter = 2;
            xlWorkSheet.SetCellValue(RowCounter, 1, "ATPI Greece Travel Marine S.A. ,31-33 Athinon Avenue, 104 47 Athens, Greece");
            xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, 15);
            RowCounter++;
            xlWorkSheet.SetCellValue(RowCounter, 1, "ATPI Ελλάς - Ταξειδιωτική Ναυτιλιακή Α.Ε., Λ.Αθηνών 31-33, 104 47, Αθήνα-ΑΦΜ 094333389 ΦΑΕ Αθηνών");
            xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, 15);
            RowCounter += 2;
            xlWorkSheet.SetCellValue(RowCounter, 1, Statement.Description);
            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, LastColumn, mStyles.xlStyleHeader);
            RowCounter++;

            xlWorkSheet.SetCellValue(RowCounter, 1, "P-T Number");
            xlWorkSheet.SetCellValue(RowCounter, 2, "Inv.Date");
            xlWorkSheet.SetCellValue(RowCounter, 3, "Vessel");
            xlWorkSheet.SetCellValue(RowCounter, 4, "Description");
            xlWorkSheet.SetCellValue(RowCounter, 5, "Pax");
            xlWorkSheet.SetCellValue(RowCounter, 6, "A/L");
            xlWorkSheet.SetCellValue(RowCounter, 7, "Routing");
            xlWorkSheet.SetCellValue(RowCounter, 8, "Flight Date");
            xlWorkSheet.SetCellValue(RowCounter, 9, "Booked By");
            xlWorkSheet.SetCellValue(RowCounter, 10, "Reason For Travel");
            xlWorkSheet.SetCellValue(RowCounter, 11, "Staff");
            xlWorkSheet.SetCellValue(RowCounter, 12, "Taxes");
            xlWorkSheet.SetCellValue(RowCounter, 13, "Face Value");
            xlWorkSheet.SetCellValue(RowCounter, 14, "Discount");
            xlWorkSheet.SetCellValue(RowCounter, 15, "Net Payable");
            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, LastColumn, mStyles.xlStyleHeader);

            for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
            {
                var Rows = mdsDataSet.Tables[0].Rows[i];

                string invoicenumber = $"{Convert.ToString(Rows[2])}";
                string clientdescription = $"{Convert.ToString(Rows[0])} {Convert.ToString(Rows[1])}";
                string vessel = $"{Convert.ToString(Rows[4])}";
                string staff = $"{Convert.ToString(Rows[12])}";
                string clientkey = $"{clientdescription}";
                string vesselkey = $"{clientdescription}|{vessel}";
                string staffkey = $"{clientdescription}|{vessel}|{staff}";
                decimal taxes = Convert.ToDecimal(Rows[13]);
                decimal facevalue = Convert.ToDecimal(Rows[14]);
                decimal discount = Convert.ToDecimal(Rows[15]);
                decimal netpayable = Convert.ToDecimal(Rows[16]);
                if (WithStaffTotal && (Staff == null || Staff.Key != staffkey))
                {
                    RowCounter = AddTotalLine(RowCounter, Staff, false);
                    Staff = new Totals("STAFF", staff, staffkey);
                }
                if (Vessel == null || Vessel.Key != vesselkey)
                {
                    RowCounter = AddTotalLine(RowCounter, Vessel, true);
                    Vessel = new Totals("VESSEL", vessel, vesselkey);
                }
                if (Client == null || Client.Key != clientkey)
                {
                    RowCounter = AddTotalLine(RowCounter, Client, true);

                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, $"Client {clientdescription}");
                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, LastColumn, mStyles.xlStyleTotalsYellow);
                    Client = new Totals("CLIENT", clientdescription, clientkey);
                }
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, invoicenumber);
                if (!string.IsNullOrEmpty(invoicenumber)) xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToDateTime(Rows[3]));
                xlWorkSheet.SetCellValue(RowCounter, 3, vessel);
                xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(Rows[5]));
                xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToInt32(Rows[6]));
                xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(Rows[7]));
                xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(Rows[8]));
                xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDateTime(Rows[9]));
                xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(Rows[10]));
                xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(Rows[11]));
                xlWorkSheet.SetCellValue(RowCounter, 11, staff);
                xlWorkSheet.SetCellValue(RowCounter, 12, taxes);
                xlWorkSheet.SetCellValue(RowCounter, 13, facevalue);
                xlWorkSheet.SetCellValue(RowCounter, 14, discount);
                xlWorkSheet.SetCellValue(RowCounter, 15, netpayable);
                if (invoicenumber == "")
                {
                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, LastColumn, mStyles.xlStyleLightSteelBlue);
                }
                Client.Add(taxes, facevalue, discount, netpayable);
                Vessel.Add(taxes, facevalue, discount, netpayable);
                if (WithStaffTotal && Staff != null) Staff.Add(taxes, facevalue, discount, netpayable);
                Statement.Add(taxes, facevalue, discount, netpayable);
            }
            RowCounter = AddTotalLine(RowCounter, Staff, false);
            RowCounter = AddTotalLine(RowCounter, Vessel, true);
            RowCounter = AddTotalLine(RowCounter, Client, true);
            RowCounter = AddTotalLine(RowCounter, Statement, true);

            xlWorkSheet.AutoFitColumn(1, LastColumn);
            xlWorkSheet.SaveAs(FileName);
            return $"File saved: {mReportTitle}";
        }
        int AddTotalLine(int rowcounter, Totals? totalline, bool highlight)
        {
            if (totalline != null)
            {
                rowcounter++;
                xlWorkSheet.SetCellValue(rowcounter, 1, $"Total {totalline.TotalType} - {totalline.Description}");
                xlWorkSheet.SetCellValue(rowcounter, 12, totalline.Taxes);
                xlWorkSheet.SetCellValue(rowcounter, 13, totalline.FaceValue);
                xlWorkSheet.SetCellValue(rowcounter, 14, totalline.Discount);
                xlWorkSheet.SetCellValue(rowcounter, 15, totalline.NetPayable);
                if (highlight)
                {
                    xlWorkSheet.SetCellStyle(rowcounter, 1, rowcounter, 15, mStyles.xlStyleTotalsYellow);
                }
                else
                {
                    xlWorkSheet.SetCellStyle(rowcounter, 1, rowcounter, 15, mStyles.xlStyleTotals);
                }
                rowcounter++;
            }

            return rowcounter;
        }
    }
    internal class Totals
    {
        internal string TotalType { get; set; }
        internal string Description { get; set; }
        internal string Key { get; set; }
        internal decimal Taxes { get; set; }
        internal decimal FaceValue { get; set; }
        internal decimal Discount { get; set; }
        internal decimal NetPayable { get; set; }
        internal Totals(string totaltype, string description, string key)
        {
            TotalType = totaltype;
            Description = description;
            Taxes = 0;
            FaceValue = 0;
            Discount = 0;
            NetPayable = 0;
            Key = key;
        }
        internal void Add(decimal taxes, decimal facevalue, decimal discount, decimal netpayable)
        {
            Taxes += taxes;
            FaceValue += facevalue;
            Discount += discount;
            NetPayable += netpayable;
        }
    }
}
