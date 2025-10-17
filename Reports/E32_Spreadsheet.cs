using PNRHistoryNext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public partial class Spreadsheets
    {
        public string E32_QCSellingPrice(PNRHistoryNext.PNRs PnrDetails, string FileName)
        {
            int RowCounter;

            try
            {
                ThisDocument.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "QCSellingPrice");
                ThisDocument.FreezePanes(1, 0);


                ThisDocument.SetColumnStyle(1, 18, xlTextStyle);
                ThisDocument.SetColumnStyle(3, 4, xlDateStyle);
                ThisDocument.SetColumnStyle(15, 18, xlIntPosStyle);

                ThisDocument.SetCellValue(1, 1, "GDS");
                ThisDocument.SetCellValue(1, 2, "PNR");
                ThisDocument.SetCellValue(1, 3, "PNR Creation Date");
                ThisDocument.SetCellValue(1, 4, "Transaction Date");
                ThisDocument.SetCellValue(1, 5, "Client Code");
                ThisDocument.SetCellValue(1, 6, "Client Name");
                ThisDocument.SetCellValue(1, 7, "Agent Group");
                ThisDocument.SetCellValue(1, 8, "Book PCC");
                ThisDocument.SetCellValue(1, 9, "Book Salesman");
                ThisDocument.SetCellValue(1, 10, "Issue OID");
                ThisDocument.SetCellValue(1, 11, "Issue Salesman");
                ThisDocument.SetCellValue(1, 12, "Product");
                ThisDocument.SetCellValue(1, 13, "Selling Price");
                ThisDocument.SetCellValue(1, 14, "Void-Dummy-Protect");
                ThisDocument.SetCellValue(1, 15, "FixedMarkupClient");
                ThisDocument.SetCellValue(1, 16, "Void");
                ThisDocument.SetCellValue(1, 17, "Refund");
                ThisDocument.SetCellValue(1, 18, "Passive");
                ThisDocument.SetCellStyle(1, 1, 1, 18, xlStyleHeader);

                RowCounter = 1;
                var gt = new PNRTotals();
                foreach (PNRHistory.PNR item in PnrDetails)
                {
                    RowCounter += 1;

                    ThisDocument.SetCellValue(RowCounter, 1, item.GDS);
                    ThisDocument.SetCellValue(RowCounter, 2, item.PNRId);
                    ThisDocument.SetCellValue(RowCounter, 3, item.PNRCreationDate);
                    ThisDocument.SetCellValue(RowCounter, 4, item.TransactionDate);
                    ThisDocument.SetCellValue(RowCounter, 5, item.ClientCode);
                    ThisDocument.SetCellValue(RowCounter, 6, item.ClientName);
                    ThisDocument.SetCellValue(RowCounter, 7, item.AgentGroup);
                    ThisDocument.SetCellValue(RowCounter, 8, item.BookPCC);
                    ThisDocument.SetCellValue(RowCounter, 9, item.BookSalesman);
                    ThisDocument.SetCellValue(RowCounter, 10, item.IssueOID);
                    ThisDocument.SetCellValue(RowCounter, 11, item.IssueSalesman);
                    ThisDocument.SetCellValue(RowCounter, 12, item.ProductName);
                    ThisDocument.SetCellValue(RowCounter, 13, item.AllSPRemarks);
                    ThisDocument.SetCellValue(RowCounter, 14, item.AllDummyRemarks);
                    ThisDocument.SetCellValue(RowCounter, 15, item.FixedMarkupClient ? 1 : 0);
                    ThisDocument.SetCellValue(RowCounter, 16, item.Void ? 1 : 0);
                    ThisDocument.SetCellValue(RowCounter, 17, item.Refund ? 1 : 0);
                    ThisDocument.SetCellValue(RowCounter, 18, item.Passive ? 1 : 0);
                }
                ThisDocument.SetCellStyle(1, 14, RowCounter, 18, xlStyleGrayItalic);
                ThisDocument.AutoFitColumn(1, 18);

                E32_Totals("Issue Agent", PnrDetails.TotsPerIssueAgent, PnrDetails.GrandTotals);
                E32_Totals("Booking Agent", PnrDetails.TotsPerBookAgent, PnrDetails.GrandTotals);
                E32_Totals("Client", PnrDetails.TotsPerClient, PnrDetails.GrandTotals);
                E32_Totals("Agent Group", PnrDetails.TotsPerAgentGroup, PnrDetails.GrandTotals);

                ThisDocument.SaveAs(FileName);
                return $"File saved: {FileName}";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
        void E32_Totals(string worksheetname, SortedDictionary<string, PNRTotals> totals, PNRTotals grandtotal)
        {
            ThisDocument.AddWorksheet(worksheetname);
            ThisDocument.FreezePanes(3, 0);
            ThisDocument.SetColumnStyle(2, 4, xlIntPosStyle);
            ThisDocument.SetColumnStyle(5, 6, xlPctPosStyle);
            ThisDocument.SetColumnStyle(7, 12, xlIntPosStyle);

            ThisDocument.SetCellValue(1, 1, "Total");
            ThisDocument.SetCellValue(1, 2, grandtotal.TotalSaleable);
            ThisDocument.SetCellValue(1, 3, grandtotal.WithSell);
            ThisDocument.SetCellValue(1, 4, grandtotal.WithoutSell);

            ThisDocument.SetCellValue(1, 5, grandtotal.WithSellPct);
            ThisDocument.SetCellValue(1, 6, grandtotal.WithoutSellPct);

            ThisDocument.SetCellValue(1, 7, grandtotal.DummyRemarks);
            ThisDocument.SetCellValue(1, 8, grandtotal.FixedMarkupClient);
            ThisDocument.SetCellValue(1, 9, grandtotal.Void);
            ThisDocument.SetCellValue(1, 10, grandtotal.Refund);
            ThisDocument.SetCellValue(1, 11, grandtotal.Passive);
            ThisDocument.SetCellValue(1, 12, grandtotal.NotAirTicket);

            ThisDocument.SetCellValue(3, 1, worksheetname);
            ThisDocument.SetCellValue(3, 2, "Total Saleable");
            ThisDocument.SetCellValue(3, 3, "With Selling Price");
            ThisDocument.SetCellValue(3, 4, "Without Selling Price");
            ThisDocument.SetCellValue(3, 5, "With Selling Price");
            ThisDocument.SetCellValue(3, 6, "Without Selling Price");
            ThisDocument.SetCellValue(3, 7, "Dummy-Protect-Visa");
            ThisDocument.SetCellValue(3, 8, "Fixed Markup Client");
            ThisDocument.SetCellValue(3, 9, "Void");
            ThisDocument.SetCellValue(3, 10, "Refund");
            ThisDocument.SetCellValue(3, 11, "Passive");
            ThisDocument.SetCellValue(3, 12, "Not Air Ticket");
            ThisDocument.SetCellStyle(3, 1, 3, 12, xlStyleHeader);
            int RowCounter = 3;
            foreach (string key in totals.Keys)
            {
                var tots = totals[key];
                RowCounter++;

                ThisDocument.SetCellValue(RowCounter, 1, key);
                ThisDocument.SetCellValue(RowCounter, 2, tots.TotalSaleable);
                ThisDocument.SetCellValue(RowCounter, 3, tots.WithSell);
                ThisDocument.SetCellValue(RowCounter, 4, tots.WithoutSell);

                ThisDocument.SetCellValue(RowCounter, 5, tots.WithSellPct);
                ThisDocument.SetCellValue(RowCounter, 6, tots.WithoutSellPct);

                ThisDocument.SetCellValue(RowCounter, 7, tots.DummyRemarks);
                ThisDocument.SetCellValue(RowCounter, 8, tots.FixedMarkupClient);
                ThisDocument.SetCellValue(RowCounter, 9, tots.Void);
                ThisDocument.SetCellValue(RowCounter, 10, tots.Refund);
                ThisDocument.SetCellValue(RowCounter, 11, tots.Passive);
                ThisDocument.SetCellValue(RowCounter, 12, tots.NotAirTicket);
            }
            ThisDocument.SetCellStyle(1, 7, RowCounter, 12, xlStyleGrayItalic);
            ThisDocument.AutoFitColumn(1, 21);
            ThisDocument.Sort(4, 1, RowCounter, 12, 2, false);
        }

    }
}
