using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public partial class Spreadsheets
    {
        public string E35_012217(GaslogReports.E35_Collection gaslogcollection, string FileName)
        {
            int RowCounter;
            int firstgrouprow = 0;
            int lastgrouprow = 0;

            try
            {
                ThisDocument.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Issue");
                E35_SetWorksheet();
                ThisDocument.AddWorksheet("Refund");
                E35_SetWorksheet();
                ThisDocument.SelectWorksheet("Issue");

                RowCounter = 1;
                GaslogReports.E35_Item PrevItem = null;
                decimal subtotal = 0;
                foreach (GaslogReports.E35_Item item in gaslogcollection.Values)
                {
                    if (item.TransactionKey.StartsWith("1") && !PrevItem.TransactionKey.StartsWith("1"))
                    {
                        if (RowCounter > 1)
                        {
                            E35_Subtotal(RowCounter, PrevItem, subtotal);
                            ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                            //ThisDocument.CollapseRows(lastgrouprow + 1);
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
                            RowCounter = E35_Subtotal(RowCounter, PrevItem, subtotal);
                            ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                            //ThisDocument.CollapseRows(lastgrouprow + 1);
                            subtotal = 0;
                        }
                        PrevItem = item;
                        RowCounter++;
                        
                        ThisDocument.SetCellStyle(RowCounter, 1, RowCounter, 32, xlStyleYellowBold);
                        firstgrouprow = RowCounter;
                    }
                    RowCounter++;

                    ThisDocument.SetCellValue(RowCounter, 1, item.RegNr);
                    ThisDocument.SetCellValue(RowCounter, 2, item.BookedBy);
                    ThisDocument.SetCellValue(RowCounter, 3, item.Office);
                    ThisDocument.SetCellValue(RowCounter, 4, item.CostCentre);
                    ThisDocument.SetCellValue(RowCounter, 5, item.ReasonForTravel);
                    ThisDocument.SetCellValue(RowCounter, 6, item.NetPayable);
                    ThisDocument.SetCellValue(RowCounter, 7, item.Routing);
                    ThisDocument.SetCellValue(RowCounter, 8, item.Passenger);
                    ThisDocument.SetCellValue(RowCounter, 9, item.Verified);                   

                    if (item.Verified != "") ThisDocument.SetCellStyle(RowCounter, 1, RowCounter, 9, xlStyleHotPink);
                    subtotal += item.NetPayable;
                    lastgrouprow = RowCounter;
                }
                if (RowCounter > 1)
                {
                    E35_Subtotal(RowCounter, PrevItem, subtotal);
                    ThisDocument.GroupRows(firstgrouprow, lastgrouprow);
                    //ThisDocument.CollapseRows(lastgrouprow + 1);
                }
                ThisDocument.AddWorksheet("Refund");
                ThisDocument.AutoFitColumn(1, 9);
                ThisDocument.SelectWorksheet("Issue");
                ThisDocument.AutoFitColumn(1, 9);
                ThisDocument.SaveAs(FileName);
                return $"File saved: {FileName}";

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
        int E35_Subtotal(int rowcounter, GaslogReports.E35_Item item, decimal subtotal)
        {
            rowcounter++;
            ThisDocument.SetCellValue(rowcounter, 2, item.BookedBy);
            ThisDocument.SetCellValue(rowcounter, 3, item.Office);
            ThisDocument.SetCellValue(rowcounter, 4, item.CostCentre);
            if (item.TransactionKey.Substring(1, 1) == "1") ThisDocument.SetCellValue(rowcounter, 5, item.ReasonForTravel);
            ThisDocument.SetCellValue(rowcounter, 6, subtotal);
            ThisDocument.SetCellStyle(rowcounter, 1, rowcounter, 9, xlStyleCyanBold);
            //rowcounter++;
            return rowcounter;
        }
        void E35_SetWorksheet()
        {
            ThisDocument.FreezePanes(1, 0);

            ThisDocument.SetColumnStyle(1, 9, xlTextStyle);
            ThisDocument.SetColumnStyle(6, 6, xlDecimalStyle);

            ThisDocument.SetCellValue(1, 1, "RegNr");
            ThisDocument.SetCellValue(1, 2, "BookedBy");
            ThisDocument.SetCellValue(1, 3, "Office");
            ThisDocument.SetCellValue(1, 4, "CostCentre");
            ThisDocument.SetCellValue(1, 5, "ReasonForTravel");
            ThisDocument.SetCellValue(1, 6, "NetPayable");
            ThisDocument.SetCellValue(1, 7, "Routing");
            ThisDocument.SetCellValue(1, 8, "Passenger");
            ThisDocument.SetCellValue(1, 9, "Verified");

            ThisDocument.SetCellStyle(1, 1, 1, 9, xlStyleHeader);
        }
    }
}
