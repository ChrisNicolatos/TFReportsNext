using GDSImportNext;
using System;

namespace ReportsNext
{
    public partial class Spreadsheets
    {
        public string E39_GDS_Import_Errors( ReportsNext.ReportsCollection report, GDSImportNext.GDSImportItems? ItemDetails, string FileName)
        {
            int RowCounter;
            try
            {
                if (ItemDetails == null)
                {
                    throw new Exception("No GDS Import Items provided");
                }
                ThisDocument.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "GDSImportErrors");
                ThisDocument.FreezePanes(1, 0);

                ThisDocument.SetColumnStyle(1, 28, xlTextStyle);
                ThisDocument.SetColumnStyle(3, 4, xlDateStyle);
                ThisDocument.SetColumnStyle(27, 27, xlDateStyle);

                ThisDocument.SetCellValue(1, 1, "Id");
                ThisDocument.SetCellValue(1, 2, "GDSDataID");
                ThisDocument.SetCellValue(1, 3, "ImportTimeStamp");
                ThisDocument.SetCellValue(1, 4, "AutoErrorTimeStamp");
                ThisDocument.SetCellValue(1, 5, "BookOfficeId");
                ThisDocument.SetCellValue(1, 6, "PNR");
                ThisDocument.SetCellValue(1, 7, "BookSalesman");
                ThisDocument.SetCellValue(1, 8, "IssueSalesman");
                ThisDocument.SetCellValue(1, 9, "Product");
                ThisDocument.SetCellValue(1, 10, "ImportStatus");
                ThisDocument.SetCellValue(1, 11, "ImportMessage");
                ThisDocument.SetCellValue(1, 12, "ImportType");
                ThisDocument.SetCellValue(1, 13, "TransactionType");
                ThisDocument.SetCellValue(1, 14, "Type");
                ThisDocument.SetCellValue(1, 15, "PNRStatus");
                ThisDocument.SetCellValue(1, 16, "PassengerSNs");
                ThisDocument.SetCellValue(1, 17, "ClientCode");
                ThisDocument.SetCellValue(1, 18, "ClientName");
                ThisDocument.SetCellValue(1, 19, "VesselRequired");
                ThisDocument.SetCellValue(1, 20, "Vessel");
                ThisDocument.SetCellValue(1, 21, "IsPassive");
                ThisDocument.SetCellValue(1, 22, "ChargeType");
                ThisDocument.SetCellValue(1, 23, "Routing");
                ThisDocument.SetCellValue(1, 24, "Supplier");
                ThisDocument.SetCellValue(1, 25, "Number");
                ThisDocument.SetCellValue(1, 26, "Passengers");
                ThisDocument.SetCellValue(1, 27, "DepartureDate");
                ThisDocument.SetCellValue(1, 28, "OwnedFileName");
                ThisDocument.SetCellValue(1, 29, "Reference");
                ThisDocument.SetCellValue(1, 30, "Value");
                ThisDocument.SetCellValue(1, 31, "Comment");
                ThisDocument.SetCellStyle(1, 1, 1, 31, xlStyleHeader);

                RowCounter = 1;
                foreach (GDSImportItem item in ItemDetails)
                {
                    if (!report.BooleanOption1 ||item.HasErrors)
                    {

                        RowCounter++;
                        ThisDocument.SetCellValue(RowCounter, 1, item.Id);
                        ThisDocument.SetCellValue(RowCounter, 2, item.GDSDataID);
                        if (item.ImportTimeStamp > DateTime.MinValue) ThisDocument.SetCellValue(RowCounter, 3, item.ImportTimeStamp);
                        if (item.AutoErrorTimeStamp > DateTime.MinValue) ThisDocument.SetCellValue(RowCounter, 4, item.AutoErrorTimeStamp);
                        ThisDocument.SetCellValue(RowCounter, 5, item.BookOfficeId);
                        ThisDocument.SetCellValue(RowCounter, 6, item.PNR);
                        ThisDocument.SetCellValue(RowCounter, 7, item.BookSalesman);
                        ThisDocument.SetCellValue(RowCounter, 8, item.IssueSalesman);
                        ThisDocument.SetCellValue(RowCounter, 9, item.Product);
                        ThisDocument.SetCellValue(RowCounter, 10, item.ImportStatus);
                        ThisDocument.SetCellValue(RowCounter, 11, item.ImportMessage);
                        ThisDocument.SetCellValue(RowCounter, 12, item.ImportType);
                        ThisDocument.SetCellValue(RowCounter, 13, item.TransactionType);
                        ThisDocument.SetCellValue(RowCounter, 14, item.Type);
                        ThisDocument.SetCellValue(RowCounter, 15, item.PNRStatus);
                        ThisDocument.SetCellValue(RowCounter, 16, item.PassengerSNs);
                        ThisDocument.SetCellValue(RowCounter, 17, item.ClientCode);
                        ThisDocument.SetCellValue(RowCounter, 18, item.ClientName);
                        ThisDocument.SetCellValue(RowCounter, 19, item.VesselRequired);
                        ThisDocument.SetCellValue(RowCounter, 20, item.Vessel);
                        ThisDocument.SetCellValue(RowCounter, 21, item.IsPassive);
                        ThisDocument.SetCellValue(RowCounter, 22, item.ChargeType);
                        ThisDocument.SetCellValue(RowCounter, 23, item.Routing);
                        ThisDocument.SetCellValue(RowCounter, 24, item.Supplier);
                        ThisDocument.SetCellValue(RowCounter, 25, item.Number);
                        ThisDocument.SetCellValue(RowCounter, 26, item.Passengers);
                        if (item.DepartureDate > DateTime.MinValue) ThisDocument.SetCellValue(RowCounter, 27, item.DepartureDate);
                        ThisDocument.SetCellValue(RowCounter, 28, item.OwnedFileName);

                        if (!report.BooleanOption1 && item.HasErrors) ThisDocument.SetCellStyle(RowCounter, 6, xlStyleYellowBold);
                        if (item.ImportStatus == "Import Failed") ThisDocument.SetCellStyle(RowCounter, 10, xlStyleHotPink);
                        if(item.ImportMessage!="") ThisDocument.SetCellStyle(RowCounter,11, xlStyleHotPink);
                        if (item.TransactionType == "Refund") this.ThisDocument.SetCellStyle(RowCounter, 1, RowCounter, 31, xlStyleRedFont);

                        if (item.TransactionType != "Refund" && item.ClientReferences.Errors.Count > 0)
                        {
                            RowCounter--;
                            foreach (GDSImportClientReference clientref in item.ClientReferences.Errors)
                            {
                                RowCounter++;
                                ThisDocument.SetCellValue(RowCounter, 29, clientref.Element);
                                ThisDocument.SetCellValue(RowCounter, 30, clientref.Value);
                                ThisDocument.SetCellValue(RowCounter, 31, clientref.Found);
                                ThisDocument.SetCellStyle(RowCounter,31, xlStyleHotPink);
                            }
                        }
                    }
                }
                ThisDocument.AutoFitColumn(1, 31);

                ThisDocument.SaveAs(FileName);
                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
    }
}
