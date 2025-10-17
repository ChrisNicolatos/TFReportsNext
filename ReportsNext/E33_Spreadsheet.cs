using PNRHistoryNext;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public partial class Spreadsheets
    {        
        
        public string E33_012212(GaslogReports.E33_Collection gaslogcollection, string FileName)
        {
            int RowCounter;
            int firstgrouprow=0;
            int lastgrouprow=0;

            try
            {
                ThisDocument.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Issue");
                E33_SetWorksheet();
                ThisDocument.AddWorksheet("Refund");
                E33_SetWorksheet();
                ThisDocument.SelectWorksheet("Issue");

                RowCounter = 1;
                GaslogReports.E33_Item PrevItem = null;
                decimal subtotal = 0;
                foreach (GaslogReports.E33_Item item in gaslogcollection.Values)
                {
                    if (item.TransactionKey.StartsWith("1") && PrevItem != null && !PrevItem.TransactionKey.StartsWith("1"))
                    {
                        if (RowCounter > 1)
                        {
                            E33_Subtotal(RowCounter, PrevItem, subtotal);
                            ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                            ThisDocument.CollapseRows(lastgrouprow +1);
                        }
                        ThisDocument.SelectWorksheet("Refund");
                        RowCounter = 1;
                        PrevItem = null;
                        subtotal = 0;
                    }

                    if (PrevItem == null || PrevItem.TransactionKey != item.TransactionKey)
                    {
                        if (RowCounter > 1)
                        {
                            RowCounter = E33_Subtotal(RowCounter, PrevItem, subtotal);                            
                            ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                            ThisDocument.CollapseRows(lastgrouprow+1 );
                            subtotal = 0;
                        }
                        PrevItem = item;
                        RowCounter++;
                        ThisDocument.SetCellValue(RowCounter, 6, $"{item.RFTHeading}");
                        if (item.TransactionKey.Substring(1, 2) == "01")
                        {
                            ThisDocument.SetCellValue(RowCounter, 7, item.Reference);
                        }
                        else
                        {
                            ThisDocument.SetCellValue(RowCounter, 8, item.Vessel);
                        }
                        ThisDocument.SetCellStyle(RowCounter, 1, RowCounter, 32, xlStyleYellowBold);
                        firstgrouprow= RowCounter;
                    }
                    RowCounter++;

                    ThisDocument.SetCellValue(RowCounter, 1, item.ClientCode);
                    ThisDocument.SetCellValue(RowCounter, 2, item.ClientName);
                    ThisDocument.SetCellValue(RowCounter, 3, item.IssueDate);
                    ThisDocument.SetCellValue(RowCounter, 4, item.TransactionType);
                    ThisDocument.SetCellValue(RowCounter, 5, item.ActionType);
                    ThisDocument.SetCellValue(RowCounter, 6, item.ReasonForTravel);
                    ThisDocument.SetCellValue(RowCounter, 7, item.Reference);
                    ThisDocument.SetCellValue(RowCounter, 8, item.Vessel);
                    ThisDocument.SetCellValue(RowCounter, 9, item.NetPayable);
                    ThisDocument.SetCellValue(RowCounter, 10, item.RegNr);
                    ThisDocument.SetCellValue(RowCounter, 11, item.ProductType);
                    ThisDocument.SetCellValue(RowCounter, 12, item.PNR);
                    ThisDocument.SetCellValue(RowCounter, 13, item.TicketNumber);
                    ThisDocument.SetCellValue(RowCounter, 14, item.Passenger);
                    ThisDocument.SetCellValue(RowCounter, 15, item.InvCode);
                    ThisDocument.SetCellValue(RowCounter, 16, item.InvSeries);
                    ThisDocument.SetCellValue(RowCounter, 17, item.InvNumber);
                    ThisDocument.SetCellValue(RowCounter, 18, item.InvoiceDate);
                    ThisDocument.SetCellValue(RowCounter, 19, item.BookedBy);
                    ThisDocument.SetCellValue(RowCounter, 20, item.Verified);
                    ThisDocument.SetCellValue(RowCounter, 21, item.Remarks);
                    ThisDocument.SetCellValue(RowCounter, 22, item.TicketingAirline);
                    ThisDocument.SetCellValue(RowCounter, 23, item.Routing);
                    ThisDocument.SetCellValue(RowCounter, 24, item.SalesPerson);
                    ThisDocument.SetCellValue(RowCounter, 25, item.IssuingAgent);
                    ThisDocument.SetCellValue(RowCounter, 26, item.CreatorAgent);
                    ThisDocument.SetCellValue(RowCounter, 27, item.DepartureDate);
                    ThisDocument.SetCellValue(RowCounter, 28, item.ArrivalDate);
                    ThisDocument.SetCellValue(RowCounter, 29, item.ConnectedDocument);
                    ThisDocument.SetCellValue(RowCounter, 30, item.DocStatusID);
                    ThisDocument.SetCellValue(RowCounter, 31, item.CancelsDoc);
                    ThisDocument.SetCellValue(RowCounter, 32, item.ServicesDescription);
                    if (item.Verified != "") ThisDocument.SetCellStyle(RowCounter, 1, RowCounter, 32, xlStyleHotPink);
                    subtotal += item.NetPayable;
                    lastgrouprow = RowCounter;
                }
                if (RowCounter > 1)
                {
                    E33_Subtotal(RowCounter, PrevItem, subtotal);                     
                    ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                    ThisDocument.CollapseRows(lastgrouprow+1);
                }
                ThisDocument.AddWorksheet("Refund");
                ThisDocument.AutoFitColumn(1, 32);
                ThisDocument.SelectWorksheet("Issue");
                ThisDocument.AutoFitColumn(1, 32);
                ThisDocument.SaveAs(FileName);
                return $"File saved: {FileName}";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
        int E33_Subtotal(int rowcounter, GaslogReports.E33_Item item, decimal subtotal)
        {
            rowcounter++;
            ThisDocument.SetCellValue(rowcounter, 6, item.RFTHeading);

            // when Reason for Travel = CA-TRAINING COMP
            if (item.TransactionKey.Substring(1, 2) == "01")
            {
                ThisDocument.SetCellValue(rowcounter, 7, item.Reference);
            }
            else
            {
                ThisDocument.SetCellValue(rowcounter, 8, item.Vessel);
            }
            ThisDocument.SetCellValue(rowcounter, 9, subtotal);
            ThisDocument.SetCellStyle(rowcounter, 1, rowcounter, 32, xlStyleCyanBold);
            //rowcounter++;
            return rowcounter;
        }
        void E33_SetWorksheet()
        {
            ThisDocument.FreezePanes(1, 0);

            ThisDocument.SetColumnStyle(1, 32, xlTextStyle);
            ThisDocument.SetColumnStyle(3, 3, xlDateStyle);
            ThisDocument.SetColumnStyle(18, 18, xlDateStyle);
            ThisDocument.SetColumnStyle(27, 28, xlDateStyle);
            ThisDocument.SetColumnStyle(9, 9, xlDecimalStyle);

            ThisDocument.SetCellValue(1, 1, "ClientCode");
            ThisDocument.SetCellValue(1, 2, "ClientName");
            ThisDocument.SetCellValue(1, 3, "IssueDate");
            ThisDocument.SetCellValue(1, 4, "TransactionType");
            ThisDocument.SetCellValue(1, 5, "ActionType");
            ThisDocument.SetCellValue(1, 6, "ReasonForTravel");
            ThisDocument.SetCellValue(1, 7, "Reference");
            ThisDocument.SetCellValue(1, 8, "Vessel");
            ThisDocument.SetCellValue(1, 9, "NetPayable");
            ThisDocument.SetCellValue(1, 10, "RegNr");
            ThisDocument.SetCellValue(1, 11, "ProductType");
            ThisDocument.SetCellValue(1, 12, "PNR");
            ThisDocument.SetCellValue(1, 13, "TicketNumber");
            ThisDocument.SetCellValue(1, 14, "Passenger");
            ThisDocument.SetCellValue(1, 15, "InvCode");
            ThisDocument.SetCellValue(1, 16, "InvSeries");
            ThisDocument.SetCellValue(1, 17, "InvNumber");
            ThisDocument.SetCellValue(1, 18, "InvoiceDate");
            ThisDocument.SetCellValue(1, 19, "BookedBy");
            ThisDocument.SetCellValue(1, 20, "Verified");
            ThisDocument.SetCellValue(1, 21, "Remarks");
            ThisDocument.SetCellValue(1, 22, "TicketingAirline");
            ThisDocument.SetCellValue(1, 23, "Routing");
            ThisDocument.SetCellValue(1, 24, "SalesPerson");
            ThisDocument.SetCellValue(1, 25, "IssuingAgent");
            ThisDocument.SetCellValue(1, 26, "CreatorAgent");
            ThisDocument.SetCellValue(1, 27, "DepartureDate");
            ThisDocument.SetCellValue(1, 28, "ArrivalDate");
            ThisDocument.SetCellValue(1, 29, "ConnectedDocument");
            ThisDocument.SetCellValue(1, 30, "DocStatusID");
            ThisDocument.SetCellValue(1, 31, "CancelsDoc");
            ThisDocument.SetCellValue(1, 32, "ServicesDescription");

            ThisDocument.SetCellStyle(1, 1, 1, 32, xlStyleHeader);
        }
    }

}
