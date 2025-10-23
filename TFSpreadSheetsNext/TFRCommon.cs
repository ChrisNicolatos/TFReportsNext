using SpreadsheetLight;
using System.Drawing;

namespace TFSpreadSheetsNext
{
    public partial class TFRCommon
    {
        const string NAMETotals = "Totals";
        const string CompanyNameEnglish = "ATPI Greece Travel Marine S.A., 31-33 Athinon Avenue, 104 47 Athens, Greece";
        const string CompanyNameGreek = "ATPI Ελλάς - Ταξειδιωτική Ναυτιλιακή Α.Ε., Λ.Αθηνών 31-33, 104 47, Αθήνα-ΑΦΜ 094333389 ΦΑΕ Αθηνών";

        public struct TotalRowProps
        {
            public int TotalRowIndex;
            public int FirstDataRow;
        }

        public struct WorkSheetProps
        {
            public string WorkSheetName;
            public int RowCountWS;
            public decimal[,] TotalsWS;
            public string PrevTotalKey1;
            public string PrevTotalKey2;
            public TotalRowProps[] GrandTotalRowsWS;
            public TotalRowProps[] AirlineTotalRowsWS;
            public TotalRowProps[] TypeTotalRowsWS;
            public int FirstRowWS;
            public int LastRowWS;
        }

        public struct WorkbookProps
        {
            public SLDocument WorkBook;
            public WorkSheetProps[] WorkSheets;
            public string WorkbookName;
            public int RowCountWB;
            public decimal[,] TotalsWB;
            public int FirstRowWB;
            public int LastRowWB;
        }


        System.Data.DataSet mdsDataSet;
       readonly string[] MonthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        string FileName;
        string mReportTitle;
        PNRHistoryNext.PNRs? pPNR;
        GaslogReports.E33_Collection? pE33_Gaslog;
        GaslogReports.E35_Collection? pE35_Gaslog;
        GDSImportNext.GDSImportItems? pGDSImport;

        ReportsNext.ReportsCollection mobjReports;
        ReportsNext.ReportsItem? mobjSelectedReport;

        SpreadsheetLight.SLDocument xlWorkSheet;
        int RowCounter;

        Styles mStyles;
        public TFRCommon(ReportsNext.ReportsItem? mobjselectedseport, ReportsNext.ReportsCollection mobjreports, System.Data.DataSet pDs, string ReportTitle, string filename)
        {
            mobjSelectedReport = mobjselectedseport;
            mobjReports = mobjreports;
            FileName = filename;
            mReportTitle = ReportTitle;
            mdsDataSet = pDs;
            xlWorkSheet = new SpreadsheetLight.SLDocument();
            mStyles = new Styles(xlWorkSheet);
        }
        public void ExportToExcel()
        {
            try
            {
                if (mobjSelectedReport == null) throw new Exception("No report selected");
                string pResponse = "";
                ReportsNext.Spreadsheets mSS2 = new();
                pPNR = [];
                switch (mobjSelectedReport.Index)
                {
                    case 7:
                        E07_ProfitPerOPSgroup();
                        break;
                    case 8:
                        E08_ProfitPerclientgroup();
                        break;
                    case 10:
                        E10_ProfitPerClientGroupExtra();
                        break;
                    case 12:
                    case 40:
                        E12_ProfitPerclientgroupwithBudgetComparison();
                        break;
                    case 18:
                        E18_AirTicketSales();
                        break;
                    case 19:
                    case 37:
                        E19_DailyProfitReportInvoicesWithTicketNumber();
                        break;
                    case 20:
                        E20_HellasConfidence();
                        break;
                    case 22:
                        E22_Euronav();
                        break;
                    case 23:
                        E23_SeaChefs();
                        break;
                    case 24:
                        E24_ProfitPerAgentTotals();
                        break;
                    case 25:
                        E25_ProfitPerAgentTransactions();
                        break;
                    case 29:
                        E29_SeaChefsDetailed();
                        break;
                    case 30:
                    case 59:
                        E30_AirTicketsFullDetails();
                        break;

                    case 32:
                        pResponse = mSS2.E32_QCSellingPrice(pPNR, FileName);
                        break;
                    case 33:
                    case 61:
                        if (pE33_Gaslog == null)
                        {
                            throw new Exception("Gaslog collection not initialized for E33 report");
                        }
                        pResponse = mSS2.E33_012212(pE33_Gaslog, FileName);
                        break;
                    case 35:
                    case 62:
                        if (pE35_Gaslog == null)
                        {
                            throw new Exception("Gaslog collection not initialized for E35 report");
                        }
                        pResponse = mSS2.E35_012217(pE35_Gaslog, FileName);
                        break;
                    case 36:
                        E36_SeaChefs_AllUnits();
                        break;
                    case 38:
                        E38_AirTicketSales();
                        break;
                    case 39:
                        pResponse = mSS2.E39_GDS_Import_Errors(mobjReports, pGDSImport, FileName);
                        break;
                    case 43:
                        E43_DailyProfitReportWithProvisionalAnalysis();
                        break;
                    case 45:
                        E45_AirTicketSalesAll();
                        break;
                    case 46:
                        E19a_ProfitReportInvoicesTotal();
                        break;
                    case 47:
                        E47_DailyProfitReportTotalsOnly();
                        break;
                    case 48:
                        E48_DailyProfitReportTotalsPerInvoice();
                        break;
                    case 51:
                        E51_Daily_Profit_Totals_per_Category();
                        break;
                    case 53:
                        E53_SeaChefs_InvoicesByDepartureDate();
                        break;
                    case 54:
                        E54_Client_Statement();
                        break;
                    case 56:
                        E56_ClientsPerGroup();
                        break;
                    case 57:
                        E57_TUI_030366();
                        break;
                    case 60:
                        E60_Report_For_Lowest_Class();
                        break;
                    case 63:
                        E63_AirTicketSalesTemenos();
                        break;
                    case 64:
                        E64_LowestClasses();
                        break;
                    case 65:
                        E65_OpsSales();
                        break;
                    case 66:
                        E66_PurchasesPerAirline();
                        break;
                    case 67:
                        E67_Columbia();
                        break;
                    case 68:
                        E68_GDSImportedPendingItems();
                        break;
                    default:
                        {
                            throw new Exception("Report not implemented in Spreadsheets");
                        }
                }

            }
            catch (Exception)
            {

                throw;
            }
        }
        public string E07_ProfitPerOPSgroup()
        {
            var moduleSignature = "TFRCommon.E07_ProfitPerOPSgroup";
            RowCounter = 0;

            try
            {
                // Profit Per OPS group
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group");
                xlWorkSheet.FreezePanes(3, 0);

                xlWorkSheet.SetColumnStyle(4, 8, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(7, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(1, 3, mStyles.xlStyleText);

                xlWorkSheet.SetCellStyle(1, 4, mStyles.xlStyleText);
                xlWorkSheet.SetCellStyle(1, 5, mStyles.xlStyleText);

                xlWorkSheet.SetCellValue(1, 4, $"{mobjReports.Date1From:dd/MM/yyyy}-{mobjReports.Date1To:dd/MM/yyyy}");

                var table = mdsDataSet.Tables[0];
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    xlWorkSheet.SetCellValue(3, j + 1, Convert.ToString(table.Columns[j].Caption));
                }
                xlWorkSheet.SetCellStyle(3, 1, 3, table.Columns.Count, mStyles.xlStyleHeader);
                RowCounter = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    RowCounter++;
                    var row = table.Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter + 3, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 4, Convert.ToDecimal(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 5, Convert.ToDecimal(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 6, Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 7, Convert.ToInt32(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 8, Convert.ToDecimal(row[7]));
                }
                xlWorkSheet.AutoFitColumn(1, 8);
                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"{moduleSignature}: {ex.Message}", ex);
            }
        }
        public string E08_ProfitPerclientgroup()
        {
            var moduleSignature = "TFRCommon.E08_ProfitPerclientgroup";
            RowCounter = 0;

            try
            {
                // Profit Per client group
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group");
                xlWorkSheet.FreezePanes(3, 0);

                xlWorkSheet.SetColumnStyle(1, 2, mStyles.xlStyleText);
                xlWorkSheet.SetColumnStyle(3, 7, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(6, mStyles.xlStyleInteger);

                xlWorkSheet.SetCellStyle(1, 4, mStyles.xlStyleText);
                xlWorkSheet.SetCellStyle(1, 5, mStyles.xlStyleText);

                xlWorkSheet.SetCellValue(1, 4, $"{mobjReports.Date1From:dd/MM/yyyy}-{mobjReports.Date1To:dd/MM/yyyy}");

                var table = mdsDataSet.Tables[0];
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    xlWorkSheet.SetCellValue(3, j + 1, Convert.ToString(table.Columns[j].Caption));
                }
                xlWorkSheet.SetCellStyle(3, 1, 3, table.Columns.Count, mStyles.xlStyleHeader);
                RowCounter = 0;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    RowCounter++;
                    var row = table.Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter + 3, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 3, Convert.ToDecimal(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 4, Convert.ToDecimal(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 5, Convert.ToDecimal(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 6, Convert.ToInt32(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter + 3, 7, Convert.ToDecimal(row[6]));
                }
                xlWorkSheet.AutoFitColumn(1, 7);
                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"{moduleSignature}: {ex.Message}", ex);
            }
        }
        public string E10_ProfitPerClientGroupExtra()
        {
            var moduleSignature = "TFRCommon.E10_ProfitPerClientGroupExtra";
            RowCounter = 0;

            try
            {
                // Profit Per client group extra
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group");
                xlWorkSheet.FreezePanes(3, 0);

                xlWorkSheet.SetColumnStyle(3, 12, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(13, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(2, mStyles.xlStyleText);

                xlWorkSheet.SetCellStyle(1, 4, mStyles.xlStyleText);
                xlWorkSheet.SetCellStyle(1, 5, mStyles.xlStyleText);

                xlWorkSheet.SetCellValue(1, 4, $"{mobjReports.Date1From:dd/MM/yyyy}-{mobjReports.Date1To:dd/MM/yyyy}");

                var table = mdsDataSet.Tables[0];
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    xlWorkSheet.SetCellValue(3, j + 1, Convert.ToString(table.Columns[j].Caption));
                }
                xlWorkSheet.SetCellStyle(3, 1, 3, table.Columns.Count, mStyles.xlStyleHeader);
                RowCounter = 3;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    RowCounter++;
                    var row = table.Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToDecimal(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDecimal(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDecimal(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToInt32(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDecimal(row[13]));

                }
                xlWorkSheet.AutoFitColumn(1, 14);
                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"{moduleSignature}: {ex.Message}", ex);
            }
        }
        public string E12_ProfitPerclientgroupwithBudgetComparison()
        {
            WorkbookProps[] pobjWorkbooks = new WorkbookProps[1];
            string pstrPrevGroupName = "";
            try
            {
                // Profit Per client group with PY
                pobjWorkbooks[0].WorkSheets = new WorkSheetProps[4];
                pobjWorkbooks[0].WorkSheets[0].WorkSheetName = "Sales";
                pobjWorkbooks[0].WorkSheets[1].WorkSheetName = "Profit";
                pobjWorkbooks[0].WorkSheets[2].WorkSheetName = "Pax";
                pobjWorkbooks[0].WorkSheets[3].WorkSheetName = "ProfitPerPax";
                for (int iSheet = 0; iSheet <= 3; iSheet++)
                {
                    pobjWorkbooks[0].WorkSheets[iSheet].RowCountWS = 0;
                    pobjWorkbooks[0].WorkSheets[iSheet].TotalsWS = new decimal[2, 26];
                    for (int ii = 0; ii < pobjWorkbooks[0].WorkSheets[iSheet].TotalsWS.GetLength(0); ii++)
                        for (int jj = 0; jj < pobjWorkbooks[0].WorkSheets[iSheet].TotalsWS.GetLength(1); jj++)
                            pobjWorkbooks[0].WorkSheets[iSheet].TotalsWS[ii, jj] = 0;
                }
                pobjWorkbooks[0].WorkBook = new SLDocument();
                pobjWorkbooks[0].WorkBook.RenameWorksheet(SLDocument.DefaultFirstSheetName, pobjWorkbooks[0].WorkSheets[0].WorkSheetName);
                E12_PrepareWorksheet(pobjWorkbooks[0].WorkBook, mobjReports);
                pobjWorkbooks[0].WorkBook.AddWorksheet(pobjWorkbooks[0].WorkSheets[1].WorkSheetName);
                E12_PrepareWorksheet(pobjWorkbooks[0].WorkBook, mobjReports);
                pobjWorkbooks[0].WorkBook.AddWorksheet(pobjWorkbooks[0].WorkSheets[2].WorkSheetName);
                E12_PrepareWorksheet(pobjWorkbooks[0].WorkBook, mobjReports);
                pobjWorkbooks[0].WorkBook.AddWorksheet(pobjWorkbooks[0].WorkSheets[3].WorkSheetName);
                E12_PrepareWorksheet(pobjWorkbooks[0].WorkBook, mobjReports);

                var a1 = mobjReports.TextEntry.Split(new[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                string pGroups = "";
                for (int i = 0; i < a1.Length; i++)
                    pGroups += a1[i].PadLeft(2, '0').Substring(a1[i].Length - 2) + "|";

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    string OpsGroup = row[0] == DBNull.Value ? string.Empty : row[0]?.ToString() ?? string.Empty;

                    if (string.IsNullOrEmpty(pGroups) || pGroups.Contains(OpsGroup.Substring(0, 2)))
                    {
                        if (pobjWorkbooks.Length == 1 || pobjWorkbooks[pobjWorkbooks.Length - 1].WorkbookName != OpsGroup)
                        {
                            pobjWorkbooks[0].WorkBook.SelectWorksheet(pobjWorkbooks[0].WorkSheets[0].WorkSheetName);
                            if (pobjWorkbooks.Length > 1)
                                E12_AddTotalsRows(ref pobjWorkbooks[0], 1, pstrPrevGroupName + " TOTAL", mobjReports.BooleanOption1);

                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[0].LastRowWS = pobjWorkbooks[0].WorkSheets[0].RowCountWS - 1;

                            Array.Resize(ref pobjWorkbooks, pobjWorkbooks.Length + 1);
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets = new WorkSheetProps[4];
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[0].WorkSheetName = "Sales";
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[1].WorkSheetName = "Profit";
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[2].WorkSheetName = "Pax";
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[3].WorkSheetName = "ProfitPerPax";
                            for (int iSheet = 0; iSheet <= 3; iSheet++)
                            {
                                pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].RowCountWS = 0;
                                pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].TotalsWS = new decimal[2, 26];
                                pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].FirstRowWS = pobjWorkbooks[0].WorkSheets[0].RowCountWS + 1;
                                pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].LastRowWS = -1;
                                for (int ii = 0; ii < pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].TotalsWS.GetLength(0); ii++)
                                    for (int jj = 0; jj < pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].TotalsWS.GetLength(1); jj++)
                                        pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[iSheet].TotalsWS[ii, jj] = 0;
                            }
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook = new SLDocument();
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkbookName = OpsGroup;
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook.RenameWorksheet(SLDocument.DefaultFirstSheetName, pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[0].WorkSheetName);
                            E12_PrepareWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook, mobjReports);
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook.AddWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[1].WorkSheetName);
                            E12_PrepareWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook, mobjReports);
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook.AddWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[2].WorkSheetName);
                            E12_PrepareWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook, mobjReports);
                            pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook.AddWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[3].WorkSheetName);
                            E12_PrepareWorksheet(pobjWorkbooks[pobjWorkbooks.Length - 1].WorkBook, mobjReports);
                        }

                        // enter new data row in master workbook
                        pstrPrevGroupName = OpsGroup;
                        pobjWorkbooks[0].WorkSheets[0].RowCountWS++;

                        for (int iWS = 0; iWS <= 2; iWS++)
                        {
                            pobjWorkbooks[0].WorkBook.SelectWorksheet(pobjWorkbooks[0].WorkSheets[iWS].WorkSheetName);
                            pobjWorkbooks[0].WorkBook.SetCellValue(pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 1, OpsGroup);
                            pobjWorkbooks[0].WorkBook.SetCellValue(pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 2, Convert.ToString(row[1]));
                            E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 3, 4, 5, Convert.ToDecimal(row[iWS + 2]), Convert.ToDecimal(row[iWS + 18]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 6, 7, 8, Convert.ToDecimal(row[iWS + 6]), Convert.ToDecimal(row[iWS + 22]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 9, 10, Convert.ToDecimal(row[iWS + 2]), Convert.ToDecimal(row[iWS + 14]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 11, 12, Convert.ToDecimal(row[iWS + 6]), Convert.ToDecimal(row[iWS + 10]), mobjReports.BooleanOption1);

                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 3] += Convert.ToDecimal(row[iWS + 2]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 3] += Convert.ToDecimal(row[iWS + 2]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 4] += Convert.ToDecimal(row[iWS + 18]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 4] += Convert.ToDecimal(row[iWS + 18]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 5] += Convert.ToDecimal(row[iWS + 6]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 5] += Convert.ToDecimal(row[iWS + 6]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 6] += Convert.ToDecimal(row[iWS + 22]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 6] += Convert.ToDecimal(row[iWS + 22]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 7] += Convert.ToDecimal(row[iWS + 14]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 7] += Convert.ToDecimal(row[iWS + 14]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[0, 8] += Convert.ToDecimal(row[iWS + 10]);
                            pobjWorkbooks[0].WorkSheets[iWS].TotalsWS[1, 8] += Convert.ToDecimal(row[iWS + 10]);
                        }

                        // Profit per Pax
                        decimal pPPPCurr = DivNums(row[3], row[4]);
                        decimal pPPPCurrBud = DivNums(row[19], row[20]);
                        decimal pPPPytd = DivNums(row[7], row[8]);
                        decimal pPPPytdBud = DivNums(row[23], row[24]);
                        decimal pPPPPrev = DivNums(row[15], row[16]);
                        decimal pPPPPrevytd = DivNums(row[11], row[12]);
                        pobjWorkbooks[0].WorkBook.SelectWorksheet(pobjWorkbooks[0].WorkSheets[3].WorkSheetName);
                        pobjWorkbooks[0].WorkBook.SetCellValue(pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 1, OpsGroup);
                        pobjWorkbooks[0].WorkBook.SetCellValue(pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 2, Convert.ToString(row[1]));
                        E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref pobjWorkbooks[0], pobjWorkbooks[0].WorkSheets[0].RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, mobjReports.BooleanOption1);

                        // enter new data row in group's workbook
                        var groupBook = pobjWorkbooks[pobjWorkbooks.Length - 1];
                        groupBook.WorkBook.SelectWorksheet(groupBook.WorkSheets[0].WorkSheetName);
                        groupBook.WorkSheets[0].RowCountWS++;
                        for (int iWS = 0; iWS <= 2; iWS++)
                        {
                            groupBook.WorkBook.SelectWorksheet(groupBook.WorkSheets[iWS].WorkSheetName);
                            groupBook.WorkBook.SetCellValue(groupBook.WorkSheets[0].RowCountWS + 3, 1, OpsGroup);
                            groupBook.WorkBook.SetCellValue(groupBook.WorkSheets[0].RowCountWS + 3, 2, Convert.ToString(row[1]));
                            E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 3, 4, 5, Convert.ToDecimal(row[iWS + 2]), Convert.ToDecimal(row[iWS + 18]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 6, 7, 8, Convert.ToDecimal(row[iWS + 6]), Convert.ToDecimal(row[iWS + 22]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 9, 10, Convert.ToDecimal(row[iWS + 2]), Convert.ToDecimal(row[iWS + 14]), mobjReports.BooleanOption1);
                            E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 11, 12, Convert.ToDecimal(row[iWS + 6]), Convert.ToDecimal(row[iWS + 10]), mobjReports.BooleanOption1);

                            groupBook.WorkSheets[iWS].TotalsWS[0, 3] += Convert.ToDecimal(row[iWS + 2]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 3] += Convert.ToDecimal(row[iWS + 2]);
                            groupBook.WorkSheets[iWS].TotalsWS[0, 4] += Convert.ToDecimal(row[iWS + 18]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 4] += Convert.ToDecimal(row[iWS + 18]);
                            groupBook.WorkSheets[iWS].TotalsWS[0, 5] += Convert.ToDecimal(row[iWS + 6]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 5] += Convert.ToDecimal(row[iWS + 6]);
                            groupBook.WorkSheets[iWS].TotalsWS[0, 6] += Convert.ToDecimal(row[iWS + 22]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 6] += Convert.ToDecimal(row[iWS + 22]);
                            groupBook.WorkSheets[iWS].TotalsWS[0, 7] += Convert.ToDecimal(row[iWS + 14]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 7] += Convert.ToDecimal(row[iWS + 14]);
                            groupBook.WorkSheets[iWS].TotalsWS[0, 8] += Convert.ToDecimal(row[iWS + 10]);
                            groupBook.WorkSheets[iWS].TotalsWS[1, 8] += Convert.ToDecimal(row[iWS + 10]);
                        }
                        // Profit per Pax for group
                        decimal gPPPCurr = DivNums(row[3], row[4]);
                        decimal gPPPCurrBud = DivNums(row[19], row[20]);
                        decimal gPPPytd = DivNums(row[7], row[8]);
                        decimal gPPPytdBud = DivNums(row[23], row[24]);
                        decimal gPPPPrev = DivNums(row[15], row[16]);
                        decimal gPPPPrevytd = DivNums(row[11], row[12]);
                        groupBook.WorkBook.SelectWorksheet(groupBook.WorkSheets[3].WorkSheetName);
                        groupBook.WorkBook.SetCellValue(groupBook.WorkSheets[0].RowCountWS + 3, 1, OpsGroup);
                        groupBook.WorkBook.SetCellValue(groupBook.WorkSheets[0].RowCountWS + 3, 2, Convert.ToString(row[1]));
                        E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 3, 4, 5, gPPPCurr, gPPPCurrBud, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 6, 7, 8, gPPPytd, gPPPytdBud, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 9, 10, gPPPCurr, gPPPPrev, mobjReports.BooleanOption1);
                        E12_AddRowValues(ref groupBook, groupBook.WorkSheets[0].RowCountWS + 3, 11, 12, gPPPytd, gPPPPrevytd, mobjReports.BooleanOption1);
                    }
                }

                E12_AddTotalsRows(ref pobjWorkbooks[0], 1, pstrPrevGroupName + "TOTAL", mobjReports.BooleanOption1);
                pobjWorkbooks[pobjWorkbooks.Length - 1].WorkSheets[0].LastRowWS = pobjWorkbooks[0].WorkSheets[0].RowCountWS - 1;
                E12_AddTotalsRows(ref pobjWorkbooks[0], 0, "Grand Total", mobjReports.BooleanOption1);

                string pAllFiles = "";
                var masterBook = pobjWorkbooks[0].WorkBook;
                for (int iWS = 0; iWS <= 3; iWS++)
                {
                    masterBook.SelectWorksheet(pobjWorkbooks[0].WorkSheets[iWS].WorkSheetName);
                    masterBook.GroupRows(pobjWorkbooks[0].WorkSheets[0].FirstRowWS + 3, pobjWorkbooks[0].WorkSheets[0].RowCountWS + 2);
                    for (int i = 1; i < pobjWorkbooks.Length; i++)
                    {
                        masterBook.GroupRows(pobjWorkbooks[i].WorkSheets[0].FirstRowWS + 3, pobjWorkbooks[i].WorkSheets[0].LastRowWS + 3);
                        masterBook.CollapseRows(pobjWorkbooks[i].WorkSheets[0].LastRowWS + 4);
                    }
                    masterBook.AutoFitColumn(0, 12);
                }
                E12_SetAreaColours(ref pobjWorkbooks[0]);
                masterBook.SaveAs(FileName);

                pAllFiles += FileName + Environment.NewLine;
                for (int iSheet = 1; iSheet < pobjWorkbooks.Length; iSheet++)
                {
                    E12_AddTotalsRows(ref pobjWorkbooks[iSheet], 0, "Total", mobjReports.BooleanOption1);
                    for (int iWS = 0; iWS <= 3; iWS++)
                    {
                        pobjWorkbooks[iSheet].WorkBook.SelectWorksheet(pobjWorkbooks[iSheet].WorkSheets[iWS].WorkSheetName);
                        pobjWorkbooks[iSheet].WorkBook.AutoFitColumn(0, 12);
                    }
                    string pFilename = FileName.Replace(".xlsx", "-" + pobjWorkbooks[iSheet].WorkbookName + ".xlsx");
                    E12_SetAreaColours(ref pobjWorkbooks[iSheet]);
                    pobjWorkbooks[iSheet].WorkBook.SaveAs(pFilename);
                    pAllFiles += pFilename + Environment.NewLine;
                }
                return pAllFiles;
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_ProfitPerclientgroupwithBudgetComparison: " + ex.Message, ex);
            }
        }
        private decimal DivNums(object Val1, object Val2)
        {
            try
            {
                decimal d1 = Convert.ToDecimal(Val1);
                decimal d2 = Convert.ToDecimal(Val2);
                if (d2 == 0) return 0;
                return d1 / d2;
            }
            catch
            {
                return 0;
            }
        }
        private void E12_AddRowValues(ref WorkbookProps pWorkBook, int Row, int FirstCell, int SecondCell, int ThirdCell, decimal Value1, decimal Value2, bool DiffAsPercentage)
        {
            try
            {
                pWorkBook.WorkBook.SetCellValue(Row, FirstCell, Value1);
                pWorkBook.WorkBook.SetCellValue(Row, SecondCell, Value2);
                decimal v1 = Convert.ToDecimal(Value1);
                decimal v2 = Convert.ToDecimal(Value2);
                if (v1 * v2 != 0)
                {
                    if (DiffAsPercentage)
                        pWorkBook.WorkBook.SetCellValue(Row, ThirdCell, v1 / v2 * 100);
                    else
                        pWorkBook.WorkBook.SetCellValue(Row, ThirdCell, v1 - v2);

                    if (v1 < v2)
                    {
                        pWorkBook.WorkBook.SetCellStyle(Row, ThirdCell, mStyles.xlStyleRedFont);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_AddRowValues: " + ex.Message, ex);
            }
        }
        private void E12_AddRowValues(ref WorkbookProps pWorkBook, int Row, int FirstCell, int SecondCell, decimal Value1, decimal Value2, bool DiffAsPercentage)
        {
            try
            {
                pWorkBook.WorkBook.SetCellValue(Row, FirstCell, Value2);
                decimal v1 = Convert.ToDecimal(Value1);
                decimal v2 = Convert.ToDecimal(Value2);
                if (v1 * v2 != 0)
                {
                    if (DiffAsPercentage)
                        pWorkBook.WorkBook.SetCellValue(Row, SecondCell, v1 / v2 * 100);
                    else
                        pWorkBook.WorkBook.SetCellValue(Row, SecondCell, v1 - v2);

                    if (v1 < v2)
                    {
                        pWorkBook.WorkBook.SetCellStyle(Row, SecondCell, mStyles.xlStyleRedFont);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_AddRowValues: " + ex.Message, ex);
            }
        }
        private void E12_SetAreaColours(ref WorkbookProps pWorkBook)
        {
            try
            {
                for (int iWS = 0; iWS <= 3; iWS++)
                {
                    pWorkBook.WorkBook.SelectWorksheet(pWorkBook.WorkSheets[iWS].WorkSheetName);
                    pWorkBook.WorkBook.SetCellStyle(1, 3, pWorkBook.WorkSheets[0].RowCountWS + 3, 5, mStyles.xlStyleGreenYellow);
                    pWorkBook.WorkBook.SetCellStyle(1, 6, pWorkBook.WorkSheets[0].RowCountWS + 3, 8, mStyles.xlStyleSpringGreen);
                    pWorkBook.WorkBook.SetCellStyle(1, 9, pWorkBook.WorkSheets[0].RowCountWS + 3, 10, mStyles.xlStylePowderBlue);
                    pWorkBook.WorkBook.SetCellStyle(1, 11, pWorkBook.WorkSheets[0].RowCountWS + 3, 12, mStyles.xlStyleSkyBlue);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_SetAreaColours: " + ex.Message, ex);
            }
        }
        private void E12_AddTotalsRows(ref WorkbookProps pWorkBook, int TotalsLevel, string TotalsDescription, bool DiffAsPercentage)
        {
            try
            {

                pWorkBook.WorkSheets[0].RowCountWS++;
                for (int iWS = 0; iWS <= 2; iWS++)
                {
                    pWorkBook.WorkBook.SelectWorksheet(pWorkBook.WorkSheets[iWS].WorkSheetName);
                    pWorkBook.WorkBook.SetCellValue(pWorkBook.WorkSheets[0].RowCountWS + 3, 1, TotalsDescription);
                    E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 3, 4, 5,
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 3],
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 4],
                        DiffAsPercentage);
                    E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 6, 7, 8,
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 5],
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 6],
                        DiffAsPercentage);
                    E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 9, 10,
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 3],
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 7],
                        DiffAsPercentage);
                    E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 11, 12,
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 5],
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, 8],
                        DiffAsPercentage);
                    pWorkBook.WorkBook.SetCellStyle(pWorkBook.WorkSheets[0].RowCountWS + 3, 1, pWorkBook.WorkSheets[0].RowCountWS + 3, 12, mStyles.xlStyleBold);
                }
                // Profit per Pax
                decimal pPPPCurr = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 3], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 3]);
                decimal pPPPCurrBud = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 4], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 4]);
                decimal pPPPytd = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 5], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 5]);
                decimal pPPPytdBud = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 6], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 6]);
                decimal pPPPPrev = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 7], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 7]);
                decimal pPPPPrevytd = DivNums(pWorkBook.WorkSheets[1].TotalsWS[TotalsLevel, 8], pWorkBook.WorkSheets[2].TotalsWS[TotalsLevel, 8]);
                pWorkBook.WorkBook.SelectWorksheet(pWorkBook.WorkSheets[3].WorkSheetName);
                pWorkBook.WorkBook.SetCellValue(pWorkBook.WorkSheets[0].RowCountWS + 3, 1, TotalsDescription);
                E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, DiffAsPercentage);
                E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, DiffAsPercentage);
                E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, DiffAsPercentage);
                E12_AddRowValues(ref pWorkBook, pWorkBook.WorkSheets[0].RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, DiffAsPercentage);
                pWorkBook.WorkBook.SetCellStyle(pWorkBook.WorkSheets[0].RowCountWS + 3, 1, pWorkBook.WorkSheets[0].RowCountWS + 3, 12, mStyles.xlStyleBold);
                for (int iWS = 0; iWS <= 2; iWS++)
                    for (int j = 0; j <= 8; j++)
                        pWorkBook.WorkSheets[iWS].TotalsWS[TotalsLevel, j] = 0;
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_AddTotalsRows: " + ex.Message, ex);
            }
        }
        private void E12_PrepareWorksheet(SLDocument xlWorkSheet, ReportsNext.ReportsCollection mReport)
        {
            try
            {
                xlWorkSheet.FreezePanes(3, 0);

                xlWorkSheet.SetColumnStyle(3, 12, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(1, 2, mStyles.xlStyleTextBold);
                xlWorkSheet.SetCellStyle(1, 3, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(1, 3, mReport.E12_FromCurr.ToString("dd/MM/yyyy") + "-" + mReport.E12_ToCurr.ToString("dd/MM/yyyy"));
                xlWorkSheet.MergeWorksheetCells(1, 3, 1, 5);
                xlWorkSheet.SetCellStyle(2, 3, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(2, 3, "Current Month");
                xlWorkSheet.MergeWorksheetCells(2, 3, 2, 5);
                xlWorkSheet.SetCellStyle(1, 6, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(1, 6, mReport.E12_FromYTD.ToString("dd/MM/yyyy") + "-" + mReport.E12_ToYTD.ToString("dd/MM/yyyy"));
                xlWorkSheet.MergeWorksheetCells(1, 6, 1, 8);
                xlWorkSheet.SetCellStyle(2, 6, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(2, 6, "Current Year To Date");
                xlWorkSheet.MergeWorksheetCells(2, 6, 2, 8);
                xlWorkSheet.SetCellStyle(1, 9, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(1, 9, mReport.E12_FromPYCurr.ToString("dd/MM/yyyy") + "-" + mReport.E12_ToPYCurr.ToString("dd/MM/yyyy"));
                xlWorkSheet.MergeWorksheetCells(1, 9, 1, 10);
                xlWorkSheet.SetCellStyle(2, 9, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(2, 9, "Prev.Year Month");
                xlWorkSheet.MergeWorksheetCells(2, 9, 2, 10);
                xlWorkSheet.SetCellStyle(1, 11, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(1, 11, mReport.E12_FromPYTD.ToString("dd/MM/yyyy") + "-" + mReport.E12_ToPYTD.ToString("dd/MM/yyyy"));
                xlWorkSheet.MergeWorksheetCells(1, 11, 1, 12);
                xlWorkSheet.SetCellStyle(2, 11, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(2, 11, "Previous Year To Date");
                xlWorkSheet.MergeWorksheetCells(2, 11, 2, 12);
                xlWorkSheet.SetCellStyle(3, 1, 3, 12, mStyles.xlStyleTextBoldCentre);
                xlWorkSheet.SetCellValue(3, 1, "Group Name");
                xlWorkSheet.SetCellValue(3, 2, "Client");
                xlWorkSheet.SetCellValue(3, 3, "Curr.Month");
                xlWorkSheet.SetCellValue(3, 4, "Budget");
                xlWorkSheet.SetCellValue(3, 5, "Comparison");
                xlWorkSheet.SetCellValue(3, 6, "YTD");
                xlWorkSheet.SetCellValue(3, 7, "Budget");
                xlWorkSheet.SetCellValue(3, 8, "Comparison");
                xlWorkSheet.SetCellValue(3, 9, "Prev.Year Month");
                xlWorkSheet.SetCellValue(3, 10, "Comparison");
                xlWorkSheet.SetCellValue(3, 11, "Prev.YTD");
                xlWorkSheet.SetCellValue(3, 12, "Comparison");
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E12_PrepareWorksheet: " + ex.Message, ex);
            }
        }
        public string E18_AirTicketSales()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 42, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(25, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(43, 45, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(15, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(36, 37, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(46, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(29, mStyles.xlStyleInteger);

                // Set headers
                xlWorkSheet.SetCellValue(1, 1, "Issue Date");
                xlWorkSheet.SetCellValue(1, 2, "Client Code");
                xlWorkSheet.SetCellValue(1, 3, "Client Name");
                if (Convert.ToBoolean(mobjReports.GetType().GetProperty("BooleanOption1")?.GetValue(mobjReports, null)))
                {
                    xlWorkSheet.SetCellValue(1, 4, "Sell");
                    xlWorkSheet.SetCellValue(1, 5, "X");
                }
                else
                {
                    xlWorkSheet.SetCellValue(1, 4, "Omit");
                    xlWorkSheet.SetCellValue(1, 5, "Void");
                }
                xlWorkSheet.SetCellValue(1, 6, "PNR");
                xlWorkSheet.SetCellValue(1, 7, "Ticket Number");
                xlWorkSheet.SetCellValue(1, 8, "Passenger");
                xlWorkSheet.SetCellValue(1, 9, "Pax Count");
                xlWorkSheet.SetCellValue(1, 10, "Product Type");
                xlWorkSheet.SetCellValue(1, 11, "Action Type");
                xlWorkSheet.SetCellValue(1, 12, "Inv Code");
                xlWorkSheet.SetCellValue(1, 13, "Inv Series");
                xlWorkSheet.SetCellValue(1, 14, "Inv Number");
                xlWorkSheet.SetCellValue(1, 15, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 16, "Vessel");
                xlWorkSheet.SetCellValue(1, 17, "Booked By");
                xlWorkSheet.SetCellValue(1, 18, "Office/Dept");
                xlWorkSheet.SetCellValue(1, 19, "Reason For Travel");
                xlWorkSheet.SetCellValue(1, 20, "Cost Centre");
                xlWorkSheet.SetCellValue(1, 21, "Requisition Number");
                xlWorkSheet.SetCellValue(1, 22, "OPT");
                xlWorkSheet.SetCellValue(1, 23, "TRID/MarineFare");
                xlWorkSheet.SetCellValue(1, 24, "Account Code");
                xlWorkSheet.SetCellValue(1, 25, "Net Payable");
                xlWorkSheet.SetCellValue(1, 26, "Verified");
                xlWorkSheet.SetCellValue(1, 27, "Remarks");
                xlWorkSheet.SetCellValue(1, 28, "Transaction Type");
                xlWorkSheet.SetCellValue(1, 29, "RegNr");
                xlWorkSheet.SetCellValue(1, 30, "Ticketing Airline");
                xlWorkSheet.SetCellValue(1, 31, "Routing");
                xlWorkSheet.SetCellValue(1, 32, "SalesPerson");
                xlWorkSheet.SetCellValue(1, 33, "Issuing Agent");
                xlWorkSheet.SetCellValue(1, 34, "Creator Agent");
                xlWorkSheet.SetCellValue(1, 35, "Reference");
                xlWorkSheet.SetCellValue(1, 36, "Departure Date");
                xlWorkSheet.SetCellValue(1, 37, "Arrival Date");
                xlWorkSheet.SetCellValue(1, 38, "Connected Document");
                xlWorkSheet.SetCellValue(1, 39, "Pax Remarks");
                xlWorkSheet.SetCellValue(1, 40, "Invoice Status");
                xlWorkSheet.SetCellValue(1, 41, "Other Services");
                xlWorkSheet.SetCellValue(1, 42, "Client Team");
                xlWorkSheet.SetCellValue(1, 43, "Sell");
                xlWorkSheet.SetCellValue(1, 44, "Buy");
                xlWorkSheet.SetCellValue(1, 45, "Profit");
                xlWorkSheet.SetCellValue(1, 46, "PaxCount+-");

                xlWorkSheet.SetCellStyle(1, 1, 1, 46, mStyles.xlStyleHeader);

                RowCounter = 1;
                var table = mdsDataSet.Tables[0];
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    bool booleanOption1 = Convert.ToBoolean(mobjReports.GetType().GetProperty("BooleanOption1")?.GetValue(mobjReports, null));
                    
                    string temp37 = row[37] == DBNull.Value ? string.Empty : row[37]?.ToString() ?? string.Empty;
                    if (!booleanOption1 ||
                        (Convert.ToString(row[3]) == "" && Convert.ToString(row[4]) == "" && temp37.Trim() == ""))
                    {
                        RowCounter++;
                        xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[0]));
                        xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                        xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                        xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                        xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                        xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                        xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                        xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7]));
                        xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToInt32(row[8]));
                        xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9]));
                        xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10]));


                        if (booleanOption1)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[25]));

                        }

                        if (Convert.ToInt32(row[13]) != 0)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11]));
                            xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[12]));
                            xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[13]));
                            xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDateTime(row[14]));
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[16]));
                        xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17]));
                        xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20]));
                        xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));
                        xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22]));
                        xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23]));
                        xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToDecimal(row[24]));
                        xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[25]));
                        xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[26]));
                        xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToString(row[27]));
                        xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToInt32(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToString(row[29]));
                        xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToString(row[30]));
                        xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToString(row[31]));
                        xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToString(row[32]));
                        xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[33]));
                        xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToString(row[34]));
                        xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToDateTime(row[35]));
                        xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToDateTime(row[36]));
                        xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToString(row[37]));
                        xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToString(row[38]));

                        if (Convert.ToInt32(row[39]) == 43)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 40, "Cancelled");
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 46, mStyles.xlStyleItalic);
                        }
                        else if (!string.IsNullOrEmpty(Convert.ToString(row[40])))
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 40, $"Cancels {Convert.ToString(row[40])}");
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 46, mStyles.xlStyleItalic);
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 41, Convert.ToString(row[41]));
                        xlWorkSheet.SetCellValue(RowCounter, 42, Convert.ToString(row[42]));

                        if (!string.IsNullOrEmpty(Convert.ToString(row[3])))
                        {
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 46, mStyles.xlStyleSandyBrown);
                        }
                        if (!string.IsNullOrEmpty(Convert.ToString(row[4])))
                        {
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 46, mStyles.xlStyleGrayItalic);
                        }
                        if (Convert.ToString(row[10]) == "Refund")
                        {
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 46, mStyles.xlStyleRedFont);
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 43, Convert.ToDecimal(row[43]));
                        xlWorkSheet.SetCellValue(RowCounter, 44, Convert.ToDecimal(row[44]));
                        xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToDecimal(row[45]));
                        xlWorkSheet.SetCellValue(RowCounter, 46, Convert.ToInt32(row[46]));

                    }
                }

                xlWorkSheet.AutoFitColumn(1, 46);
                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E18_AirTicketSales: " + ex.Message, ex);
            }
        }
        public string E19_DailyProfitReportInvoicesWithTicketNumber()
        {
            // Field numbers from SQL (0-based)
            int DBITEMAirProfitPerPax = 26;
            int DBITEMSerProfitPerPax = 45;
            int DBITEMTotProfitPerPax = 64;


            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Report Invoices");
                xlWorkSheet.FreezePanes(3, 8);

                xlWorkSheet.SetColumnStyle(9, 65, mStyles.xlStyleDecimalBlankZero);

                xlWorkSheet.SetColumnStyle(4, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(26, mStyles.xlStyleDecimalBlankZero);
                xlWorkSheet.SetColumnStyle(45, mStyles.xlStyleDecimalBlankZero);
                xlWorkSheet.SetColumnStyle(64, mStyles.xlStyleDecimalBlankZero);

                xlWorkSheet.SetCellValue(1, 9, $"{mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                xlWorkSheet.SetCellValue(2, 9, "Air Tickets");
                xlWorkSheet.SetCellValue(2, 28, "Other Services");
                xlWorkSheet.SetCellValue(2, 47, "Total");

                xlWorkSheet.SetCellValue(3, 1, "Client Code");
                xlWorkSheet.SetCellValue(3, 2, "Client Name");
                xlWorkSheet.SetCellValue(3, 3, "Vessel");
                xlWorkSheet.SetCellValue(3, 4, "Invoice Date");
                xlWorkSheet.SetCellValue(3, 5, "Invoice Type");
                xlWorkSheet.SetCellValue(3, 6, "Invoice Number");
                xlWorkSheet.SetCellValue(3, 7, "Airline");
                xlWorkSheet.SetCellValue(3, 8, "Ticket Number");

                for (int i = 9; i <= 48; i += 19)
                {
                    xlWorkSheet.SetCellValue(3, i, "Face value");
                    xlWorkSheet.SetCellValue(3, i + 1, "Taxes");
                    xlWorkSheet.SetCellValue(3, i + 2, "Commission");
                    xlWorkSheet.SetCellValue(3, i + 3, "Discount");
                    xlWorkSheet.SetCellValue(3, i + 4, "Cancellation Fee");
                    xlWorkSheet.SetCellValue(3, i + 5, "TF");
                    xlWorkSheet.SetCellValue(3, i + 6, "Net Payable");
                    xlWorkSheet.SetCellValue(3, i + 7, "Net Buy");
                    xlWorkSheet.SetCellValue(3, i + 8, "IW5");
                    xlWorkSheet.SetCellValue(3, i + 9, "IW6");
                    xlWorkSheet.SetCellValue(3, i + 10, "IW7");
                    xlWorkSheet.SetCellValue(3, i + 11, "IW8");
                    xlWorkSheet.SetCellValue(3, i + 12, "IW9");
                    xlWorkSheet.SetCellValue(3, i + 13, "IW11");
                    xlWorkSheet.SetCellValue(3, i + 14, "IW10");
                    xlWorkSheet.SetCellValue(3, i + 15, "IW");
                    xlWorkSheet.SetCellValue(3, i + 16, "Profit");
                    xlWorkSheet.SetCellValue(3, i + 17, "Pax");
                    xlWorkSheet.SetCellValue(3, i + 18, "Profit Per Pax");
                }
                xlWorkSheet.SetCellValue(3, 66, "Omit");
                xlWorkSheet.SetCellValue(3, 67, "Client Team");
                xlWorkSheet.SetCellValue(3, 68, "Customer Group");
                xlWorkSheet.SetCellStyle(1, 9, 3, 68, mStyles.xlStyleHeader);

                xlWorkSheet.MergeWorksheetCells(1, 9, 1, 65);
                xlWorkSheet.MergeWorksheetCells(2, 9, 2, 27);
                xlWorkSheet.MergeWorksheetCells(2, 28, 2, 46);
                xlWorkSheet.MergeWorksheetCells(2, 47, 2, 65);

                RowCounter = 3;
                decimal[] pTotals = new decimal[65];
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    RowCounter++;
                    for (int j = 0; j <= 64; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                        xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                        xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                        xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDateTime(row[3]));
                        xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                        xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                        xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                        xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7]));
                        xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[8]));
                        xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[9]));
                        xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[10]));
                        xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[11]));
                        xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[12]));
                        xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDecimal(row[13]));
                        xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDecimal(row[14]));
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[16]));
                        xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToDecimal(row[17]));
                        xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToDecimal(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToDecimal(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToDecimal(row[20]));
                        xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToDecimal(row[21]));
                        xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToDecimal(row[22]));
                        xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToDecimal(row[23]));
                        xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToDecimal(row[24]));
                        xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToInt32(row[25]));
                        xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToDecimal(row[26]));
                        xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToDecimal(row[27]));
                        xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDecimal(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToDecimal(row[29]));
                        xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToDecimal(row[30]));
                        xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToDecimal(row[31]));
                        xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToDecimal(row[32]));
                        xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToDecimal(row[33]));
                        xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToDecimal(row[34]));
                        xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToDecimal(row[35]));
                        xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToDecimal(row[36]));
                        xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToDecimal(row[37]));
                        xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToDecimal(row[38]));
                        xlWorkSheet.SetCellValue(RowCounter, 40, Convert.ToDecimal(row[39]));
                        xlWorkSheet.SetCellValue(RowCounter, 41, Convert.ToDecimal(row[40]));
                        xlWorkSheet.SetCellValue(RowCounter, 42, Convert.ToDecimal(row[41]));
                        xlWorkSheet.SetCellValue(RowCounter, 43, Convert.ToDecimal(row[42]));
                        xlWorkSheet.SetCellValue(RowCounter, 44, Convert.ToDecimal(row[43]));
                        xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToInt32(row[44]));
                        xlWorkSheet.SetCellValue(RowCounter, 46, Convert.ToDecimal(row[45]));
                        xlWorkSheet.SetCellValue(RowCounter, 47, Convert.ToDecimal(row[46]));
                        xlWorkSheet.SetCellValue(RowCounter, 48, Convert.ToDecimal(row[47]));
                        xlWorkSheet.SetCellValue(RowCounter, 49, Convert.ToDecimal(row[48]));
                        xlWorkSheet.SetCellValue(RowCounter, 50, Convert.ToDecimal(row[49]));
                        xlWorkSheet.SetCellValue(RowCounter, 51, Convert.ToDecimal(row[50]));
                        xlWorkSheet.SetCellValue(RowCounter, 52, Convert.ToDecimal(row[51]));
                        xlWorkSheet.SetCellValue(RowCounter, 53, Convert.ToDecimal(row[52]));
                        xlWorkSheet.SetCellValue(RowCounter, 54, Convert.ToDecimal(row[53]));
                        xlWorkSheet.SetCellValue(RowCounter, 55, Convert.ToDecimal(row[54]));
                        xlWorkSheet.SetCellValue(RowCounter, 56, Convert.ToDecimal(row[55]));
                        xlWorkSheet.SetCellValue(RowCounter, 57, Convert.ToDecimal(row[56]));
                        xlWorkSheet.SetCellValue(RowCounter, 58, Convert.ToDecimal(row[57]));
                        xlWorkSheet.SetCellValue(RowCounter, 59, Convert.ToDecimal(row[58]));
                        xlWorkSheet.SetCellValue(RowCounter, 60, Convert.ToDecimal(row[59]));
                        xlWorkSheet.SetCellValue(RowCounter, 61, Convert.ToDecimal(row[60]));
                        xlWorkSheet.SetCellValue(RowCounter, 62, Convert.ToDecimal(row[61]));
                        xlWorkSheet.SetCellValue(RowCounter, 63, Convert.ToDecimal(row[62]));
                        xlWorkSheet.SetCellValue(RowCounter, 64, Convert.ToInt32(row[63]));
                        xlWorkSheet.SetCellValue(RowCounter, 65, Convert.ToDecimal(row[64]));

                        if (j > 7)
                        {
                            pTotals[j] += Convert.ToDecimal(row[j]);
                        }
                    }

                    if (Convert.ToBoolean(row[65]))
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 66, "Omit");
                        xlWorkSheet.SetCellStyle(RowCounter, 5, RowCounter, 6, mStyles.xlStyleSandyBrown);
                    }
                    xlWorkSheet.SetCellValue(RowCounter, 67, Convert.ToString(row[67]));
                    xlWorkSheet.SetCellValue(RowCounter, 68, Convert.ToString(row[68]));

                }

                pTotals[DBITEMAirProfitPerPax] = pTotals[DBITEMAirProfitPerPax - 1] != 0
                    ? pTotals[DBITEMAirProfitPerPax - 2] / pTotals[DBITEMAirProfitPerPax - 1]
                    : 0;
                pTotals[DBITEMSerProfitPerPax] = pTotals[DBITEMSerProfitPerPax - 1] != 0
                    ? pTotals[DBITEMSerProfitPerPax - 2] / pTotals[DBITEMSerProfitPerPax - 1]
                    : 0;
                pTotals[DBITEMTotProfitPerPax] = pTotals[DBITEMTotProfitPerPax - 1] != 0
                    ? pTotals[DBITEMTotProfitPerPax - 2] / pTotals[DBITEMTotProfitPerPax - 1]
                    : 0;

                xlWorkSheet.SetCellValue(RowCounter + 1, 2, "TOTAL");
                for (int j = 8; j <= 64; j++)
                {
                    xlWorkSheet.SetCellValue(RowCounter + 1, j + 1, pTotals[j]);
                }

                for (int iNeg = 0; iNeg < mdsDataSet.Tables[0].Rows.Count; iNeg++)
                {
                    var row = mdsDataSet.Tables[0].Rows[iNeg];
                    for (int iProf = DBITEMAirProfitPerPax; iProf <= DBITEMTotProfitPerPax; iProf += 19)
                    {
                        if (Convert.ToDecimal(row[iProf]) < pTotals[iProf])
                        {
                            xlWorkSheet.SetCellStyle(iNeg + 4, iProf + 1, mStyles.xlStyleRedFont);
                        }
                    }
                }

                xlWorkSheet.SetCellStyle(RowCounter + 1, 9, RowCounter + 1, 67, mStyles.xlStyleHeader);

                xlWorkSheet.SetCellStyle(3, DBITEMAirProfitPerPax + 1, RowCounter + 1, DBITEMAirProfitPerPax + 1, mStyles.xlStyleDecimalYellow);
                xlWorkSheet.SetCellStyle(3, DBITEMSerProfitPerPax + 1, RowCounter + 1, DBITEMSerProfitPerPax + 1, mStyles.xlStyleDecimalYellow);
                xlWorkSheet.SetCellStyle(3, DBITEMTotProfitPerPax + 1, RowCounter + 1, DBITEMTotProfitPerPax + 1, mStyles.xlStyleDecimalYellow);

                xlWorkSheet.SetColumnStyle(26, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(45, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(64, mStyles.xlStyleInteger);

                xlWorkSheet.AutoFitColumn(0, 68);

                xlWorkSheet.SetCellStyle(3, 9, RowCounter + 1, 14, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetCellStyle(3, 28, RowCounter + 1, 33, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetCellStyle(3, 47, RowCounter + 1, 52, mStyles.xlStyleGrayItalic);
                xlWorkSheet.GroupColumns(9, 14);
                xlWorkSheet.CollapseColumns(15);
                xlWorkSheet.GroupColumns(28, 33);
                xlWorkSheet.CollapseColumns(34);
                xlWorkSheet.GroupColumns(47, 52);
                xlWorkSheet.CollapseColumns(53);

                xlWorkSheet.SetColumnStyle(17, 23, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetColumnStyle(36, 42, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetColumnStyle(55, 61, mStyles.xlStyleGrayItalic);
                xlWorkSheet.GroupColumns(17, 23);
                xlWorkSheet.CollapseColumns(24);
                xlWorkSheet.GroupColumns(36, 42);
                xlWorkSheet.CollapseColumns(43);
                xlWorkSheet.GroupColumns(55, 61);
                xlWorkSheet.CollapseColumns(62);

                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("TFRCommon.E19_DailyProfitReportInvoicesWithTicketNumber: " + ex.Message, ex);
            }
        }
        public string E19a_ProfitReportInvoicesTotal()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Totals");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 14, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(2, 14, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(5, 5, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(9, 9, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(13, 13, mStyles.xlStyleInteger);

                xlWorkSheet.SetCellValue(1, 1, "Client Group");
                xlWorkSheet.SetCellValue(1, 2, "Net Payable Invoices");
                xlWorkSheet.SetCellValue(1, 3, "Net Buy Invoices");
                xlWorkSheet.SetCellValue(1, 4, "Profit Invoices");
                xlWorkSheet.SetCellValue(1, 5, "Pax Invoices");
                xlWorkSheet.SetCellValue(1, 6, "Net Payable CN");
                xlWorkSheet.SetCellValue(1, 7, "Net Buy CN");
                xlWorkSheet.SetCellValue(1, 8, "Profit CN");
                xlWorkSheet.SetCellValue(1, 9, "Pax CN");
                xlWorkSheet.SetCellValue(1, 10, "Net Payable");
                xlWorkSheet.SetCellValue(1, 11, "Net Buy");
                xlWorkSheet.SetCellValue(1, 12, "Profit");
                xlWorkSheet.SetCellValue(1, 13, "Pax");
                xlWorkSheet.SetCellValue(1, 14, "% CN Pax/Total Pax");

                xlWorkSheet.SetCellStyle(1, 1, 1, 14, mStyles.xlStyleHeader);

                RowCounter = 1;
                var table = mdsDataSet.Tables[0];
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    RowCounter++;
                    var row = table.Rows[i];
                    // Use Convert.ToXXX to prevent type errors
                    xlWorkSheet.SetCellValue(RowCounter, 1, row[0]?.ToString());
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToDecimal(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToDecimal(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDecimal(row[1]) - Convert.ToDecimal(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDecimal(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[4]) - Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[1]) + Convert.ToDecimal(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[2]) + Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, (Convert.ToDecimal(row[1]) - Convert.ToDecimal(row[2])) + (Convert.ToDecimal(row[4]) - Convert.ToDecimal(row[5])));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[3]) + Convert.ToDecimal(row[6]));
                    if (Convert.ToDecimal(row[3]) != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 14, -Convert.ToDecimal(row[6]) / Convert.ToDecimal(row[3]) * 100);
                    }
                }
                xlWorkSheet.Sort(2, 1, RowCounter, 14, 12, false);
                xlWorkSheet.AutoFitColumn(1, 14);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E19aProfitReportInvoicesTotal: " + ex.Message, ex);
            }
        }
        public string E20_HellasConfidence()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Client Invoices");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 22, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(16, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(14, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleInteger);

                // Set header row
                xlWorkSheet.SetCellValue(1, 1, "Issue Date");
                xlWorkSheet.SetCellValue(1, 2, "Client Code");
                xlWorkSheet.SetCellValue(1, 3, "Client Name");
                xlWorkSheet.SetCellValue(1, 4, "PNR");
                xlWorkSheet.SetCellValue(1, 5, "Ticket Number");
                xlWorkSheet.SetCellValue(1, 6, "Passenger");
                xlWorkSheet.SetCellValue(1, 7, "Details");
                xlWorkSheet.SetCellValue(1, 8, "Pax Count");
                xlWorkSheet.SetCellValue(1, 9, "Product Type");
                xlWorkSheet.SetCellValue(1, 10, "Action Type");
                xlWorkSheet.SetCellValue(1, 11, "Inv Code");
                xlWorkSheet.SetCellValue(1, 12, "Inv Series");
                xlWorkSheet.SetCellValue(1, 13, "Inv Number");
                xlWorkSheet.SetCellValue(1, 14, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 15, "Vessel");
                xlWorkSheet.SetCellValue(1, 16, "Net Payable");
                xlWorkSheet.SetCellValue(1, 17, "Transaction Type");
                xlWorkSheet.SetCellValue(1, 18, "Booked By");
                xlWorkSheet.SetCellValue(1, 19, "Office/Dept");
                xlWorkSheet.SetCellValue(1, 20, "Reason For Travel");
                xlWorkSheet.SetCellValue(1, 21, "Cost Centre");
                xlWorkSheet.SetCellValue(1, 22, "Requisition Number");

                xlWorkSheet.SetCellStyle(1, 1, 1, 22, mStyles.xlStyleHeader);

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToInt32(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToInt32(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDateTime(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[15]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[16]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17]));
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20]));
                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));

                    if (Convert.ToString(row[9]) == "Refund")
                    {
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 22, mStyles.xlStyleRedFont);
                    }
                }

                xlWorkSheet.AutoFitColumn(1, 22);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E20_HellasConfidence: " + ex.Message, ex);
            }
        }
        public string E22_Euronav()
        {

            RowCounter = 0;
            int pSheet = 0;
            int columnShift = mobjReports.BooleanOption1 ? 1 : 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "INV");
                pSheet += 1;
                while (pSheet < 3)
                {
                    xlWorkSheet.FreezePanes(1, 0);

                    xlWorkSheet.SetColumnStyle(1, 18 + columnShift, mStyles.xlStyleText);

                    xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDecimal);

                    xlWorkSheet.SetColumnStyle(3, mStyles.xlStyleDate);
                    xlWorkSheet.SetColumnStyle(7, mStyles.xlStyleDate);

                    xlWorkSheet.SetCellValue(1, 1, "Invoice Type");
                    xlWorkSheet.SetCellValue(1, 2, "Invoice Number");
                    xlWorkSheet.SetCellValue(1, 3, "Invoice Date");
                    xlWorkSheet.SetCellValue(1, 4, "Vessel Name");
                    xlWorkSheet.SetCellValue(1, 5, "Routing");
                    xlWorkSheet.SetCellValue(1, 6, "Destination");
                    xlWorkSheet.SetCellValue(1, 7, "Departure Date");
                    xlWorkSheet.SetCellValue(1, 8, "Turnover");
                    if (columnShift == 1)
                    {
                        xlWorkSheet.SetCellValue(1, 9, "Currency");
                    }
                    xlWorkSheet.SetCellValue(1, 9 + columnShift, "BookedBy");
                    xlWorkSheet.SetCellValue(1, 10 + columnShift, "CPDepartment");
                    xlWorkSheet.SetCellValue(1, 11 + columnShift, "CID Code");
                    xlWorkSheet.SetCellValue(1, 12 + columnShift, "Trip ID");
                    xlWorkSheet.SetCellValue(1, 13 + columnShift, "PassengerName");
                    xlWorkSheet.SetCellValue(1, 14 + columnShift, "Passenger ID");
                    xlWorkSheet.SetCellValue(1, 15 + columnShift, "Nationality");
                    xlWorkSheet.SetCellValue(1, 16 + columnShift, "ClientCode");
                    xlWorkSheet.SetCellValue(1, 17 + columnShift, "PNRID");
                    xlWorkSheet.SetCellValue(1, 18 + columnShift, "ConnectedDocument");
                    xlWorkSheet.SetCellStyle(1, 1, 1, 18 + columnShift, mStyles.xlStyleHeader);

                    RowCounter = 1;
                    for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                    {
                        var row = mdsDataSet.Tables[0].Rows[i];
                        // Use Convert.ToString/ToInt32/ToDecimal/ToDateTime to avoid type errors
                        if ((pSheet == 1 && Convert.ToString(row[10]) != "Refund") ||
                            (pSheet == 2 && Convert.ToString(row[10]) == "Refund"))
                        {
                            if (Convert.ToString(row[3]) == "" && Convert.ToString(row[4]) == "" && Convert.ToString(row[25]) == "")
                            {
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 1, $"{Convert.ToString(row[11])} {Convert.ToString(row[12])}");
                                xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToInt32(row[13]));
                                xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToDateTime(row[14]));
                                xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[15]));
                                var routing = Convert.ToString(row[30]);
                                xlWorkSheet.SetCellValue(RowCounter, 6, "");
                                xlWorkSheet.SetCellValue(RowCounter, 7, "");
                                if (!string.IsNullOrEmpty(routing))
                                {
                                    xlWorkSheet.SetCellValue(RowCounter, 5, routing);
                                    if (Convert.ToString(row[27]) == "AIR" && routing.Length > 3)
                                    {
                                        xlWorkSheet.SetCellValue(RowCounter, 6, routing.Substring(routing.Length - 3));
                                        xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDateTime(row[35]));
                                    }
                                }

                                if (columnShift == 0)
                                {
                                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[24]));
                                }
                                else
                                {
                                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[39]));
                                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[40]));
                                }
                                xlWorkSheet.SetCellValue(RowCounter, 9 + columnShift, Convert.ToString(row[16]));
                                xlWorkSheet.SetCellValue(RowCounter, 10 + columnShift, Convert.ToString(row[17]));
                                xlWorkSheet.SetCellValue(RowCounter, 11 + columnShift, Convert.ToString(row[23]));
                                xlWorkSheet.SetCellValue(RowCounter, 12 + columnShift, Convert.ToString(row[22]));
                                xlWorkSheet.SetCellValue(RowCounter, 13 + columnShift, Convert.ToString(row[7]));
                                xlWorkSheet.SetCellValue(RowCounter, 14 + columnShift, Convert.ToString(row[38]));
                                xlWorkSheet.SetCellValue(RowCounter, 15 + columnShift, Convert.ToString(row[19]));
                                xlWorkSheet.SetCellValue(RowCounter, 16 + columnShift, Convert.ToString(row[1]));
                                xlWorkSheet.SetCellValue(RowCounter, 17 + columnShift, Convert.ToString(row[5]));
                                xlWorkSheet.SetCellValue(RowCounter, 18 + columnShift, Convert.ToString(row[37]));

                                if (Convert.ToString(row[3]) != "")
                                {
                                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 18 + columnShift, mStyles.xlStyleSandyBrown);
                                }
                                if (Convert.ToString(row[4]) != "")
                                {
                                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 18 + columnShift, mStyles.xlStyleGrayItalic);
                                }
                                if (Convert.ToString(row[11]) == "Refund")
                                {
                                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 18 + columnShift, mStyles.xlStyleRedFont);
                                }
                            }
                        }
                    }
                    xlWorkSheet.AutoFitColumn(1, 18 + columnShift);
                    xlWorkSheet.AddWorksheet("CNS");
                    pSheet += 1;
                }
                xlWorkSheet.SelectWorksheet("INV");
                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E22_Euronav: {ex.Message}", ex);
            }
        }
        public string E23_SeaChefs()
        {

            int pLastColumn = 0;
            RowCounter = 0;
            int xlHeaderID = 0;
            string pInvNumber = "";
            int pInvRow = 0;
            decimal pInvSum = 0;
            var pBookedBy = new List<string>();
            int pWorksheetCount = 0;

            try
            {
                // Build list of BookedBy if BooleanOption1, else just one worksheet
                if (mobjReports.BooleanOption1)
                {
                    for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                    {
                        var row = mdsDataSet.Tables[0].Rows[i];
                        string bookedBy = row[32] == DBNull.Value ? string.Empty : row[32]?.ToString() ?? string.Empty;
                        if (!pBookedBy.Contains(bookedBy))
                            pBookedBy.Add(bookedBy);
                    }
                }
                else
                {
                    pBookedBy.Add("");
                }

                pBookedBy.Sort();
                foreach (var bby in pBookedBy)
                {
                    pLastColumn = 143;
                    RowCounter = 0;
                    xlHeaderID = 0;
                    pInvNumber = "";
                    pInvRow = 0;
                    pInvSum = 0;
                    pWorksheetCount++;

                    // Worksheet setup
                    if (pWorksheetCount > 1)
                    {
                        xlWorkSheet.AddWorksheet(bby);
                    }
                    else
                    {
                        if (!mobjReports.BooleanOption1)
                            xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices");
                        else
                            xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, bby);
                    }

                    xlWorkSheet.FreezePanes(8, 0);

                    xlWorkSheet.SetColumnStyle(1, pLastColumn, mStyles.xlStyleText);

                    xlWorkSheet.SetColumnStyle(10, mStyles.xlStyleDecimal);
                    xlWorkSheet.SetColumnStyle(80, mStyles.xlStyleDecimal);

                    xlWorkSheet.SetColumnStyle(11, mStyles.xlStyleDate);

                    xlWorkSheet.SetCellValue(3, 3, "Create Invoices");
                    xlWorkSheet.SetCellStyle(3, 3, mStyles.xlStyleTitle);

                    // Header row
                    xlWorkSheet.SetCellValue(8, 1, "");
                    xlWorkSheet.SetCellValue(8, 2, "");
                    xlWorkSheet.SetCellValue(8, 3, "Changed");
                    xlWorkSheet.SetCellValue(8, 4, "Row Status");
                    xlWorkSheet.SetCellValue(8, 5, "*Invoice Header Identifier");
                    xlWorkSheet.SetCellValue(8, 6, "*Business Unit");
                    xlWorkSheet.SetCellValue(8, 7, "Import Set");
                    xlWorkSheet.SetCellValue(8, 8, "*Invoice Number");
                    xlWorkSheet.SetCellValue(8, 9, "*Invoice Currency");
                    xlWorkSheet.SetCellValue(8, 10, "*Invoice Amount");
                    xlWorkSheet.SetCellValue(8, 11, "*Invoice Date");
                    xlWorkSheet.SetCellValue(8, 12, "**Supplier[..]");
                    xlWorkSheet.SetCellValue(8, 13, "**Supplier Number");
                    xlWorkSheet.SetCellValue(8, 14, "*Supplier Site[..]");
                    xlWorkSheet.SetCellValue(8, 15, "Payment Currency");
                    xlWorkSheet.SetCellValue(8, 16, "Invoice Type");
                    xlWorkSheet.SetCellValue(8, 17, "Description");
                    xlWorkSheet.SetCellValue(8, 18, "Legal Entity Name[..]");
                    xlWorkSheet.SetCellValue(8, 19, "Payment Terms");
                    xlWorkSheet.SetCellValue(8, 20, "Terms Date");
                    xlWorkSheet.SetCellValue(8, 21, "Goods Received Date");
                    xlWorkSheet.SetCellValue(8, 22, "Invoice Received Date");
                    xlWorkSheet.SetCellValue(8, 23, "Accounting Date");
                    xlWorkSheet.SetCellValue(8, 24, "Budget Date");
                    xlWorkSheet.SetCellValue(8, 25, "Invoice Includes Prepayment");
                    xlWorkSheet.SetCellValue(8, 26, "Prepayment Number[..]");
                    xlWorkSheet.SetCellValue(8, 27, "Prepayment Line");
                    xlWorkSheet.SetCellValue(8, 28, "Prepayment Application Amount");
                    xlWorkSheet.SetCellValue(8, 29, "Prepayment Accounting Date");
                    xlWorkSheet.SetCellValue(8, 30, "Payment Method");
                    xlWorkSheet.SetCellValue(8, 31, "Pay Group");
                    xlWorkSheet.SetCellValue(8, 32, "Pay Alone");
                    xlWorkSheet.SetCellValue(8, 33, "Conversion Rate Type");
                    xlWorkSheet.SetCellValue(8, 34, "Conversion Date");
                    xlWorkSheet.SetCellValue(8, 35, "Conversion Rate");
                    xlWorkSheet.SetCellValue(8, 36, "Payment Cross-Conversion Rate Type");
                    xlWorkSheet.SetCellValue(8, 37, "Payment Cross-Conversion Date");
                    xlWorkSheet.SetCellValue(8, 38, "Payment Cross-Conversion Rate");
                    xlWorkSheet.SetCellValue(8, 39, "Discountable Amount");
                    xlWorkSheet.SetCellValue(8, 40, "Liability Distribution[..]");
                    xlWorkSheet.SetCellValue(8, 41, "Remit-to Supplier");
                    xlWorkSheet.SetCellValue(8, 42, "Remit-to Supplier Number");
                    xlWorkSheet.SetCellValue(8, 43, "Remit-to Address Name");
                    xlWorkSheet.SetCellValue(8, 44, "Remit-to Account Number");
                    xlWorkSheet.SetCellValue(8, 45, "Document Category");
                    xlWorkSheet.SetCellValue(8, 46, "Voucher Number");
                    xlWorkSheet.SetCellValue(8, 47, "First-Party Tax Registration Number");
                    xlWorkSheet.SetCellValue(8, 48, "Supplier Tax Registration Number");
                    xlWorkSheet.SetCellValue(8, 49, "Requester");
                    xlWorkSheet.SetCellValue(8, 50, "Delivery Channel");
                    xlWorkSheet.SetCellValue(8, 51, "Bank Charge Bearer");
                    xlWorkSheet.SetCellValue(8, 52, "Settlement Priority");
                    xlWorkSheet.SetCellValue(8, 53, "Unique Remittance Identifier");
                    xlWorkSheet.SetCellValue(8, 54, "Unique Remittance Identifier Check Digit");
                    xlWorkSheet.SetCellValue(8, 55, "Payment Reason");
                    xlWorkSheet.SetCellValue(8, 56, "Payment Reason Comments");
                    xlWorkSheet.SetCellValue(8, 57, "Remittance Message 1");
                    xlWorkSheet.SetCellValue(8, 58, "Remittance Message 2");
                    xlWorkSheet.SetCellValue(8, 59, "Remittance Message 3");
                    xlWorkSheet.SetCellValue(8, 60, "Taxation Country");
                    xlWorkSheet.SetCellValue(8, 61, "Document Subtype");
                    xlWorkSheet.SetCellValue(8, 62, "Invoice Internal Sequence");
                    xlWorkSheet.SetCellValue(8, 63, "Tax Related Invoice");
                    xlWorkSheet.SetCellValue(8, 64, "Supplier Tax Invoice Number");
                    xlWorkSheet.SetCellValue(8, 65, "Internal Recording Date");
                    xlWorkSheet.SetCellValue(8, 66, "Supplier Tax Invoice Date");
                    xlWorkSheet.SetCellValue(8, 67, "Supplier Tax Invoice Conversion Rate");
                    xlWorkSheet.SetCellValue(8, 68, "Customs Location Code");
                    xlWorkSheet.SetCellValue(8, 69, "Correction Year");
                    xlWorkSheet.SetCellValue(8, 70, "Correction Period");
                    xlWorkSheet.SetCellValue(8, 71, "Tax Control Amount");
                    xlWorkSheet.SetCellValue(8, 72, "URL Attachment");
                    xlWorkSheet.SetCellValue(8, 73, "Context Value[..]");
                    xlWorkSheet.SetCellValue(8, 74, "Additional Information[..]");
                    xlWorkSheet.SetCellValue(8, 75, "Regional Context Value [..]");
                    xlWorkSheet.SetCellValue(8, 76, "Regional Information [..]");
                    xlWorkSheet.SetCellValue(8, 77, "[..]");
                    xlWorkSheet.SetCellValue(8, 78, "Line");
                    xlWorkSheet.SetCellValue(8, 79, "*Type");
                    xlWorkSheet.SetCellValue(8, 80, "*Amount");
                    xlWorkSheet.SetCellValue(8, 81, "Invoiced Quantity");
                    xlWorkSheet.SetCellValue(8, 82, "Unit Price");
                    xlWorkSheet.SetCellValue(8, 83, "UOM");
                    xlWorkSheet.SetCellValue(8, 84, "Description");
                    xlWorkSheet.SetCellValue(8, 85, "Purchase Order[..]");
                    xlWorkSheet.SetCellValue(8, 86, "Purchase Order Line[..]");
                    xlWorkSheet.SetCellValue(8, 87, "Purchase Order Schedule[..]");
                    xlWorkSheet.SetCellValue(8, 88, "Purchase Order Distribution[..]");
                    xlWorkSheet.SetCellValue(8, 89, "Item Description");
                    xlWorkSheet.SetCellValue(8, 90, "Receipt[..]");
                    xlWorkSheet.SetCellValue(8, 91, "Receipt Line[..]");
                    xlWorkSheet.SetCellValue(8, 92, "Consumption Advice[..]");
                    xlWorkSheet.SetCellValue(8, 93, "Consumption Advice Line Number[..]");
                    xlWorkSheet.SetCellValue(8, 94, "Landed Cost Enabled");
                    xlWorkSheet.SetCellValue(8, 95, "Final Match");
                    xlWorkSheet.SetCellValue(8, 96, "Distribution Combination[..]");
                    xlWorkSheet.SetCellValue(8, 97, "Distribution Set[..]");
                    xlWorkSheet.SetCellValue(8, 98, "Accounting Date");
                    xlWorkSheet.SetCellValue(8, 99, "Overlay Account Segment");
                    xlWorkSheet.SetCellValue(8, 100, "Overlay Primary Balancing Segment");
                    xlWorkSheet.SetCellValue(8, 101, "Overlay Cost Center Segment");
                    xlWorkSheet.SetCellValue(8, 102, "Budget Date");
                    xlWorkSheet.SetCellValue(8, 103, "Tax Classification Code");
                    xlWorkSheet.SetCellValue(8, 104, "Ship-to Location[..]");
                    xlWorkSheet.SetCellValue(8, 105, "Ship-from Location[..]");
                    xlWorkSheet.SetCellValue(8, 106, "Location of Final Discharge[..]");
                    xlWorkSheet.SetCellValue(8, 107, "Regime Code");
                    xlWorkSheet.SetCellValue(8, 108, "Tax Code");
                    xlWorkSheet.SetCellValue(8, 109, "Jurisdiction Code");
                    xlWorkSheet.SetCellValue(8, 110, "Tax Status Code");
                    xlWorkSheet.SetCellValue(8, 111, "Rate Code");
                    xlWorkSheet.SetCellValue(8, 112, "Rate");
                    xlWorkSheet.SetCellValue(8, 113, "Withholding Tax Group");
                    xlWorkSheet.SetCellValue(8, 114, "Income Tax Type");
                    xlWorkSheet.SetCellValue(8, 115, "Income Tax Region");
                    xlWorkSheet.SetCellValue(8, 116, "Prorate Across All Item Lines");
                    xlWorkSheet.SetCellValue(8, 117, "Line Group Number");
                    xlWorkSheet.SetCellValue(8, 118, "Transaction Business Category");
                    xlWorkSheet.SetCellValue(8, 119, "Product Fiscal Classification");
                    xlWorkSheet.SetCellValue(8, 120, "Intended Use");
                    xlWorkSheet.SetCellValue(8, 121, "User-Defined Fiscal Classification");
                    xlWorkSheet.SetCellValue(8, 122, "Product Type");
                    xlWorkSheet.SetCellValue(8, 123, "Assessable Value");
                    xlWorkSheet.SetCellValue(8, 124, "Product Category");
                    xlWorkSheet.SetCellValue(8, 125, "Tax Control Amount");
                    xlWorkSheet.SetCellValue(8, 126, "Statistical Quantity");
                    xlWorkSheet.SetCellValue(8, 127, "Deferred Accounting Option");
                    xlWorkSheet.SetCellValue(8, 128, "Multiperiod Accounting Start Date");
                    xlWorkSheet.SetCellValue(8, 129, "Multiperiod Accounting End Date");
                    xlWorkSheet.SetCellValue(8, 130, "Track as Asset");
                    xlWorkSheet.SetCellValue(8, 131, "Serial Number");
                    xlWorkSheet.SetCellValue(8, 132, "Book");
                    xlWorkSheet.SetCellValue(8, 133, "Asset Category");
                    xlWorkSheet.SetCellValue(8, 134, "Manufacturer");
                    xlWorkSheet.SetCellValue(8, 135, "Model");
                    xlWorkSheet.SetCellValue(8, 136, "Requester");
                    xlWorkSheet.SetCellValue(8, 137, "Item ID");
                    xlWorkSheet.SetCellValue(8, 138, "Context Value[..]");
                    xlWorkSheet.SetCellValue(8, 139, "Additional Information[..]");
                    xlWorkSheet.SetCellValue(8, 140, "Project Information[..]");
                    xlWorkSheet.SetCellValue(8, 141, "Fiscal Charge Type");
                    xlWorkSheet.SetCellValue(8, 142, "Multiperiod Accounting Accrual Account[..]");
                    xlWorkSheet.SetCellValue(8, 143, "Key");

                    xlWorkSheet.SetCellStyle(8, 1, 8, pLastColumn, mStyles.xlStyleHeader);
                    xlWorkSheet.SetRowHeight(8, 56.25);

                    RowCounter = 8;
                    if (mdsDataSet.Tables[0].Rows.Count > 0)
                    {
                        var firstRow = mdsDataSet.Tables[0].Rows[0];
                        pInvNumber = firstRow[4] == DBNull.Value ? string.Empty : firstRow[4]?.ToString() ?? string.Empty;
                    }
                    pInvRow = RowCounter + 1;
                    pInvSum = 0;
                    xlHeaderID = 1;

                    for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                    {
                        var row = mdsDataSet.Tables[0].Rows[i];
                        if (!mobjReports.BooleanOption1 || Convert.ToString(row[32]) == bby)
                        {
                            RowCounter++;
                            if (Convert.ToString(row[4]) == pInvNumber)
                            {
                                pInvSum += Convert.ToDecimal(row[16]);
                            }
                            else
                            {
                                xlHeaderID++;
                                xlWorkSheet.SetCellValue(pInvRow, 10, pInvSum);
                                for (int k = pInvRow + 1; k < RowCounter; k++)
                                {
                                    xlWorkSheet.SetCellStyle(k, 6, mStyles.xlStyleGrey);
                                    xlWorkSheet.SetCellStyle(k, 10, mStyles.xlStyleGrey);
                                    xlWorkSheet.SetCellStyle(k, 47, mStyles.xlStyleGrey);
                                    xlWorkSheet.SetCellValue(k, 6, "");
                                    xlWorkSheet.SetCellValue(k, 10, "");
                                    xlWorkSheet.SetCellValue(k, 47, "");
                                }
                                pInvSum = Convert.ToDecimal(row[16]);
                                pInvNumber = row[4] == DBNull.Value ? string.Empty : row[4]?.ToString() ?? string.Empty;
                                pInvRow = RowCounter;
                            }

                            xlWorkSheet.SetCellValue(RowCounter, 5, xlHeaderID);
                            xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[3]));
                            xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToInt32(row[4]));
                            xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[5]));

                            xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDateTime(row[7]));
                            xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[8]));
                            xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[9]));
                            xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[10]));
                            xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[11]));
                            xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[12]));
                            xlWorkSheet.SetCellValue(RowCounter, 47, Convert.ToString(row[13]));

                            xlWorkSheet.SetCellValue(RowCounter, 78, Convert.ToInt32(row[14]));
                            xlWorkSheet.SetCellValue(RowCounter, 79, Convert.ToString(row[15]));
                            xlWorkSheet.SetCellValue(RowCounter, 80, Convert.ToDecimal(row[16]));

                            xlWorkSheet.SetCellValue(RowCounter, 84, Convert.ToString(row[17]));
                            xlWorkSheet.SetCellValue(RowCounter, 96, Convert.ToString(row[19]));
                            xlWorkSheet.SetCellValue(RowCounter, 103, Convert.ToString(row[44]));
                        }
                    }

                    if (!string.IsNullOrEmpty(pInvNumber) && pInvSum != 0)
                    {
                        xlWorkSheet.SetCellValue(pInvRow, 10, pInvSum);
                        for (int k = pInvRow + 1; k <= RowCounter; k++)
                        {
                            xlWorkSheet.SetCellStyle(k, 6, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellStyle(k, 10, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellStyle(k, 47, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellValue(k, 6, "");
                            xlWorkSheet.SetCellValue(k, 10, "");
                            xlWorkSheet.SetCellValue(k, 47, "");
                        }
                    }

                    xlWorkSheet.AutoFitColumn(1, pLastColumn);
                    xlWorkSheet.HideColumn(7);
                    xlWorkSheet.HideColumn(18, 46);
                    xlWorkSheet.HideColumn(48, 77);
                    xlWorkSheet.HideColumn(81, 83);
                    xlWorkSheet.HideColumn(85, 95);
                    xlWorkSheet.HideColumn(97, 102);
                    xlWorkSheet.HideColumn(104, 143);
                }

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E23_SeaChefs: " + ex.Message, ex);
            }
        }
        public string E24_ProfitPerAgentTotals()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 10, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(6, 8, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(10, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);

                // Header
                xlWorkSheet.SetCellValue(1, 1, "Group Name");
                xlWorkSheet.SetCellValue(1, 2, "Creator Agent");
                xlWorkSheet.SetCellValue(1, 3, "Sales Person");
                xlWorkSheet.SetCellValue(1, 4, "Client Code");
                xlWorkSheet.SetCellValue(1, 5, "Client Name");
                xlWorkSheet.SetCellValue(1, 6, "Sales");
                xlWorkSheet.SetCellValue(1, 7, "Cost");
                xlWorkSheet.SetCellValue(1, 8, "Profit");
                xlWorkSheet.SetCellValue(1, 9, "Pax");
                xlWorkSheet.SetCellValue(1, 10, "Profit per Pax");

                xlWorkSheet.SetCellStyle(1, 1, 1, 10, mStyles.xlStyleHeader);

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToInt32(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[9]));

                }

                xlWorkSheet.AutoFitColumn(1, 10);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E24_ProfitPerAgentTotals: " + ex.Message, ex);
            }
        }
        public string E25_ProfitPerAgentTransactions()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 12, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(8, 10, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(12, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(11, mStyles.xlStyleInteger);

                xlWorkSheet.SetColumnStyle(6, mStyles.xlStyleDate);

                // Header
                xlWorkSheet.SetCellValue(1, 1, "Group Name");
                xlWorkSheet.SetCellValue(1, 2, "Creator Agent");
                xlWorkSheet.SetCellValue(1, 3, "Sales Person");
                xlWorkSheet.SetCellValue(1, 4, "Client Code");
                xlWorkSheet.SetCellValue(1, 5, "Client Name");
                xlWorkSheet.SetCellValue(1, 6, "Issue Date");
                xlWorkSheet.SetCellValue(1, 7, "Doc Number");
                xlWorkSheet.SetCellValue(1, 8, "Sales");
                xlWorkSheet.SetCellValue(1, 9, "Cost");
                xlWorkSheet.SetCellValue(1, 10, "Profit");
                xlWorkSheet.SetCellValue(1, 11, "Pax");
                xlWorkSheet.SetCellValue(1, 12, "Profit per Pax");

                xlWorkSheet.SetCellStyle(1, 1, 1, 12, mStyles.xlStyleHeader);

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDateTime(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToInt32(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[11]));


                }

                xlWorkSheet.AutoFitColumn(1, 12);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E25_ProfitPerAgentTransactions: " + ex.Message, ex);
            }
        }
        public string E29_SeaChefsDetailed()
        {

            int pLastColumn = 23;
            RowCounter = 0;
            int xlHeaderID = 0;
            string pInvNumber = "";
            int pInvRow = 0;
            decimal pInvSum = 0;

            try
            {
                // Determine last column based on BooleanOption1
                if (mobjReports.BooleanOption1)
                    pLastColumn = 44;

                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices");
                xlWorkSheet.FreezePanes(8, 0);

                xlWorkSheet.SetColumnStyle(1, pLastColumn, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(7, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(17, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(3, 3, "Create Invoices");
                xlWorkSheet.SetCellStyle(3, 3, mStyles.xlStyleTitle);

                // Header rows
                xlWorkSheet.SetCellValue(7, 1, "");
                xlWorkSheet.SetCellValue(7, 2, "");
                xlWorkSheet.SetCellValue(7, 3, "");
                xlWorkSheet.SetCellValue(7, 4, "FIXED");
                xlWorkSheet.SetCellValue(7, 5, "Invoice");
                xlWorkSheet.SetCellValue(7, 6, "FIXED");
                xlWorkSheet.SetCellValue(7, 7, "");
                xlWorkSheet.SetCellValue(7, 8, "Batch Date");
                xlWorkSheet.SetCellValue(7, 9, "FIXED");
                xlWorkSheet.SetCellValue(7, 10, "FIXED");
                xlWorkSheet.SetCellValue(7, 11, "FIXED");
                xlWorkSheet.SetCellValue(7, 12, "FIXED");
                xlWorkSheet.SetCellValue(7, 13, "Standard/Credit Memo");
                xlWorkSheet.SetCellValue(7, 14, "FIXED");
                xlWorkSheet.SetCellValue(7, 15, "FIXED");
                xlWorkSheet.SetCellValue(7, 16, "FIXED");
                xlWorkSheet.SetCellValue(7, 17, "Payable");
                xlWorkSheet.SetCellValue(7, 18, "AL / Booked By / Routing / Traveller / Dep date/TRID");
                xlWorkSheet.SetCellValue(7, 19, "Flight / P & I / Accommodation");
                xlWorkSheet.SetCellValue(7, 20, "Cost Centre");
                xlWorkSheet.SetCellValue(7, 21, "Vessel");
                xlWorkSheet.SetCellValue(7, 22, "Client Code");
                xlWorkSheet.SetCellValue(7, 23, "Client Name");

                xlWorkSheet.SetCellStyle(7, 4, 7, 4, mStyles.xlStyleFixed);
                xlWorkSheet.SetCellStyle(7, 6, 7, 6, mStyles.xlStyleFixed);
                xlWorkSheet.SetCellStyle(7, 9, 7, 12, mStyles.xlStyleFixed);
                xlWorkSheet.SetCellStyle(7, 14, 7, 16, mStyles.xlStyleFixed);
                xlWorkSheet.SetCellStyle(7, 5, 7, 5, mStyles.xlStyleHeaderNotes);
                xlWorkSheet.SetCellStyle(7, 8, 7, 8, mStyles.xlStyleHeaderNotes);
                xlWorkSheet.SetCellStyle(7, 13, 7, 13, mStyles.xlStyleHeaderNotes);
                xlWorkSheet.SetCellStyle(7, 17, 7, pLastColumn, mStyles.xlStyleHeaderNotes);

                xlWorkSheet.SetCellValue(8, 1, "Changed");
                xlWorkSheet.SetCellValue(8, 2, "Row Status");
                xlWorkSheet.SetCellValue(8, 3, "Invoice Header Identifier");
                xlWorkSheet.SetCellValue(8, 4, "Business Unit");
                xlWorkSheet.SetCellValue(8, 5, "Invoice Number");
                xlWorkSheet.SetCellValue(8, 6, "Invoice Currency");
                xlWorkSheet.SetCellValue(8, 7, "Invoice Amount");
                xlWorkSheet.SetCellValue(8, 8, "Invoice Date");
                xlWorkSheet.SetCellValue(8, 9, "Supplier");
                xlWorkSheet.SetCellValue(8, 10, "Supplier Number");
                xlWorkSheet.SetCellValue(8, 11, "Supplier Site");
                xlWorkSheet.SetCellValue(8, 12, "Payment Currency");
                xlWorkSheet.SetCellValue(8, 13, "Invoice Type");
                xlWorkSheet.SetCellValue(8, 14, "First-Party Tax Registration Number");
                xlWorkSheet.SetCellValue(8, 15, "Line");
                xlWorkSheet.SetCellValue(8, 16, "Type");
                xlWorkSheet.SetCellValue(8, 17, "Amount");
                xlWorkSheet.SetCellValue(8, 18, "Description");
                xlWorkSheet.SetCellValue(8, 19, "Expense Type");
                xlWorkSheet.SetCellValue(8, 20, "Distribution Combination");
                xlWorkSheet.SetCellValue(8, 21, "Vessel");
                xlWorkSheet.SetCellValue(8, 22, "Client Code");
                xlWorkSheet.SetCellValue(8, 23, "Client Name");

                if (mobjReports.BooleanOption1)
                {
                    xlWorkSheet.SetCellValue(8, 24, "Invoice Code");
                    xlWorkSheet.SetCellValue(8, 25, "Invoice Series");
                    xlWorkSheet.SetCellValue(8, 26, "PNR");
                    xlWorkSheet.SetCellValue(8, 27, "Airline Code");
                    xlWorkSheet.SetCellValue(8, 28, "Ticket Number");
                    xlWorkSheet.SetCellValue(8, 29, "Pax Count");
                    xlWorkSheet.SetCellValue(8, 30, "Product Type");
                    xlWorkSheet.SetCellValue(8, 31, "Product Type LT");
                    xlWorkSheet.SetCellValue(8, 32, "Action Type");
                    xlWorkSheet.SetCellValue(8, 33, "Id");
                    xlWorkSheet.SetCellValue(8, 34, "Office");
                    xlWorkSheet.SetCellValue(8, 35, "Reason For Travel");
                    xlWorkSheet.SetCellValue(8, 36, "Requisition Number");
                    xlWorkSheet.SetCellValue(8, 37, "OPT");
                    xlWorkSheet.SetCellValue(8, 38, "TRID");
                    xlWorkSheet.SetCellValue(8, 39, "Verified");
                    xlWorkSheet.SetCellValue(8, 40, "Remarks");
                    xlWorkSheet.SetCellValue(8, 41, "RegNr");
                    xlWorkSheet.SetCellValue(8, 42, "Sales Person");
                    xlWorkSheet.SetCellValue(8, 43, "Issuing Agent");
                    xlWorkSheet.SetCellValue(8, 44, "Creator Agent");
                }

                xlWorkSheet.SetCellStyle(8, 1, 8, pLastColumn, mStyles.xlStyleHeader);
                xlWorkSheet.SetRowHeight(8, 56.25);

                RowCounter = 8;
                var firstRow = mdsDataSet.Tables[0].Rows[0];
                pInvNumber = firstRow[4] == DBNull.Value ? string.Empty : firstRow[4]?.ToString() ?? string.Empty;
                pInvRow = RowCounter + 1;
                pInvSum = 0;
                xlHeaderID = 1;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    RowCounter++;

                    if (Convert.ToString(row[4]) == pInvNumber)
                    {
                        pInvSum += Convert.ToDecimal(row[16]);
                    }
                    else
                    {
                        xlHeaderID++;
                        xlWorkSheet.SetCellValue(pInvRow, 7, pInvSum);
                        for (int k = pInvRow + 1; k < RowCounter; k++)
                        {
                            xlWorkSheet.SetCellStyle(k, 4, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellStyle(k, 7, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellStyle(k, 14, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellValue(k, 4, "");
                            xlWorkSheet.SetCellValue(k, 7, "");
                            xlWorkSheet.SetCellValue(k, 14, "");
                        }
                        pInvSum = Convert.ToDecimal(row[16]);
                        pInvNumber = row[4] == DBNull.Value ? string.Empty : row[4]?.ToString() ?? string.Empty;
                        pInvRow = RowCounter;
                    }

                    xlWorkSheet.SetCellValue(RowCounter, 3, xlHeaderID);
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToInt32(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));

                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDateTime(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToInt32(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[16]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17]));
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20]));
                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));
                    xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22]));
                    if (mobjReports.BooleanOption1)
                    {

                        xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23]));
                        xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[24]));
                        xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[25]));
                        xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[26]));
                        xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToString(row[27]));
                        xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToInt32(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToString(row[29]));
                        xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToInt32(row[30]));
                        xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToString(row[31]));
                        xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToString(row[32]));
                        xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[33]));
                        xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToString(row[34]));
                        xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToString(row[35]));
                        xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToString(row[36]));
                        xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToString(row[37]));
                        xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToString(row[38]));
                        xlWorkSheet.SetCellValue(RowCounter, 40, Convert.ToString(row[39]));
                        xlWorkSheet.SetCellValue(RowCounter, 41, Convert.ToString(row[40]));
                        xlWorkSheet.SetCellValue(RowCounter, 42, Convert.ToInt32(row[41]));
                        xlWorkSheet.SetCellValue(RowCounter, 43, Convert.ToString(row[42]));
                        xlWorkSheet.SetCellValue(RowCounter, 44, Convert.ToString(row[43]));
                        xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToString(row[44]));
                    }
                }

                if (!string.IsNullOrEmpty(pInvNumber) && pInvSum != 0)
                {
                    xlWorkSheet.SetCellValue(pInvRow, 7, pInvSum);
                    for (int k = pInvRow + 1; k <= RowCounter; k++)
                    {
                        xlWorkSheet.SetCellStyle(k, 4, mStyles.xlStyleGrey);
                        xlWorkSheet.SetCellStyle(k, 7, mStyles.xlStyleGrey);
                        xlWorkSheet.SetCellStyle(k, 14, mStyles.xlStyleGrey);
                        xlWorkSheet.SetCellValue(k, 4, "");
                        xlWorkSheet.SetCellValue(k, 7, "");
                        xlWorkSheet.SetCellValue(k, 14, "");
                    }
                }

                xlWorkSheet.AutoFitColumn(1, pLastColumn);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E29_SeaChefsDetailed: " + ex.Message, ex);
            }
        }
        public string E30_AirTicketsFullDetails()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales");
                xlWorkSheet.FreezePanes(1, 0);




                xlWorkSheet.SetColumnStyle(1, 67, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(25, 36, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(25, 26, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(15, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(47, 48, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(40, mStyles.xlStyleInteger);

                // Header row
                string[] headers = {
            "Issue Date", "Client Code", "Client Name", "Omit", "Void", "PNR", "Ticket Number", "Passenger", "Pax Count", "Product Type", "Action Type",
            "Inv Code", "Inv Series", "Inv Number", "Invoice Date", "Vessel", "Booked By", "Office/Dept", "Reason For Travel", "Cost Centre",
            "Requisition Number", "OPT", "TRID/MarineFare", "Account Code", "FC", "NetPayable FC", "Net Payable", "Net Purchase", "Face Value", "Taxes",
            "Discount", "Commission", "Cancellation Fee", "FV Extra", "Tax Extra", "Service Fee", "Verified", "Remarks", "Transaction Type", "RegNr",
            "Ticketing Airline", "Routing", "SalesPerson", "Issuing Agent", "Creator Agent", "Reference", "Departure Date", "Arrival Date", "Connected Document",
            "Pax Remarks", "Rank", "Nationality", "REF1", "REF2", "REF3", "REF4", "REF5", "REF6", "REF7", "REF8", "REF9", "REF10", "REF11", "REF12", "REF13",
            "REF14", "REF15", "Net Remit"
        };
                for (int h = 0; h < headers.Length; h++)
                    xlWorkSheet.SetCellValue(1, h + 1, headers[h]);
                xlWorkSheet.SetCellStyle(1, 1, 1, 68, mStyles.xlStyleHeader);

                // Data rows
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];

                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToInt32(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10]));


                    // Inv fields (11-14) only if InvNumber != 0
                    if (Convert.ToInt32(row[13]) != 0)
                    {
                        for (int j = 11; j <= 14; j++)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11]));
                            xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[12]));
                            xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[13]));
                            xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDateTime(row[14]));
                        }
                    }

                    // Columns 15-23
                    for (int j = 15; j <= 23; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[16]));
                        xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17]));
                        xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20]));
                        xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));
                        xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22]));
                        xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23]));

                    }


                    // FC and NetPayable FC
                    xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[66]));
                    xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToDecimal(row[65]));

                    if (Convert.ToInt32(row[67]) == 1)
                        xlWorkSheet.SetCellStyle(RowCounter, 25, RowCounter, 26, mStyles.xlStyleLightGrayItalic);

                    xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToDecimal(row[24]));
                    xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToDecimal(row[25]));
                    xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDecimal(row[26]));
                    xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToDecimal(row[27]));
                    xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToDecimal(row[28]));
                    xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToDecimal(row[29]));
                    xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToDecimal(row[30]));
                    xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToDecimal(row[31]));
                    xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToDecimal(row[32]));
                    xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToDecimal(row[33]));
                    xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToString(row[34]));
                    xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToString(row[35]));
                    xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToString(row[36]));
                    xlWorkSheet.SetCellValue(RowCounter, 40, Convert.ToInt32(row[37]));
                    xlWorkSheet.SetCellValue(RowCounter, 41, Convert.ToString(row[38]));
                    xlWorkSheet.SetCellValue(RowCounter, 42, Convert.ToString(row[39]));
                    xlWorkSheet.SetCellValue(RowCounter, 43, Convert.ToString(row[40]));
                    xlWorkSheet.SetCellValue(RowCounter, 44, Convert.ToString(row[41]));
                    xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToString(row[42]));
                    xlWorkSheet.SetCellValue(RowCounter, 46, Convert.ToString(row[43]));
                    xlWorkSheet.SetCellValue(RowCounter, 47, Convert.ToDateTime(row[44]));
                    xlWorkSheet.SetCellValue(RowCounter, 48, Convert.ToDateTime(row[45]));
                    xlWorkSheet.SetCellValue(RowCounter, 49, Convert.ToString(row[46]));
                    xlWorkSheet.SetCellValue(RowCounter, 50, Convert.ToString(row[47]));
                    xlWorkSheet.SetCellValue(RowCounter, 51, Convert.ToString(row[48]));
                    xlWorkSheet.SetCellValue(RowCounter, 52, Convert.ToString(row[49]));
                    xlWorkSheet.SetCellValue(RowCounter, 53, Convert.ToString(row[50]));
                    xlWorkSheet.SetCellValue(RowCounter, 54, Convert.ToString(row[51]));
                    xlWorkSheet.SetCellValue(RowCounter, 55, Convert.ToString(row[52]));
                    xlWorkSheet.SetCellValue(RowCounter, 56, Convert.ToString(row[53]));
                    xlWorkSheet.SetCellValue(RowCounter, 57, Convert.ToString(row[54]));
                    xlWorkSheet.SetCellValue(RowCounter, 58, Convert.ToString(row[55]));
                    xlWorkSheet.SetCellValue(RowCounter, 59, Convert.ToString(row[56]));
                    xlWorkSheet.SetCellValue(RowCounter, 60, Convert.ToString(row[57]));
                    xlWorkSheet.SetCellValue(RowCounter, 61, Convert.ToString(row[58]));
                    xlWorkSheet.SetCellValue(RowCounter, 62, Convert.ToString(row[59]));
                    xlWorkSheet.SetCellValue(RowCounter, 63, Convert.ToString(row[60]));
                    xlWorkSheet.SetCellValue(RowCounter, 64, Convert.ToString(row[61]));
                    xlWorkSheet.SetCellValue(RowCounter, 65, Convert.ToString(row[62]));
                    xlWorkSheet.SetCellValue(RowCounter, 66, Convert.ToString(row[63]));
                    xlWorkSheet.SetCellValue(RowCounter, 67, Convert.ToString(row[64]));


                    // Net Remit
                    xlWorkSheet.SetCellValue(RowCounter, 68, Convert.ToString(row[68]));

                    // Style for Omit/Void/Refund
                    if (!string.IsNullOrEmpty(Convert.ToString(row[3])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 68, mStyles.xlStyleSandyBrown);
                    if (!string.IsNullOrEmpty(Convert.ToString(row[4])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 68, mStyles.xlStyleGrayItalic);
                    if (Convert.ToString(row[10]) == "Refund")
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 68, mStyles.xlStyleRedFont);
                }

                xlWorkSheet.AutoFitColumn(1, 68);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E30_AirTicketsFullDetails: {ex.Message}", ex);
            }
        }
        public string E36_SeaChefs_AllUnits()
        {

            RowCounter = 0;
            int xlHeaderID = 0;
            string pInvNumber = "";
            int pInvRow = 0;
            decimal pInvSum = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 32, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(10, 12, mStyles.xlStyleDate);

                // Header row
                xlWorkSheet.SetCellValue(1, 1, "Business Unit");
                xlWorkSheet.SetCellValue(1, 2, "Vessel");
                xlWorkSheet.SetCellValue(1, 3, "ClientCode");
                xlWorkSheet.SetCellValue(1, 4, "ClientName");
                xlWorkSheet.SetCellValue(1, 5, "InvCode");
                xlWorkSheet.SetCellValue(1, 6, "InvSeries");
                xlWorkSheet.SetCellValue(1, 7, "Invoice Number");
                xlWorkSheet.SetCellValue(1, 8, "Invoice Currency");
                xlWorkSheet.SetCellValue(1, 9, "Invoice Amount");
                xlWorkSheet.SetCellValue(1, 10, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 11, "FromDate");
                xlWorkSheet.SetCellValue(1, 12, "ToDate");
                xlWorkSheet.SetCellValue(1, 13, "Line");
                xlWorkSheet.SetCellValue(1, 14, "Description");
                xlWorkSheet.SetCellValue(1, 15, "Distribution Combination");
                xlWorkSheet.SetCellValue(1, 16, "PNR");
                xlWorkSheet.SetCellValue(1, 17, "TicketNumber");
                xlWorkSheet.SetCellValue(1, 18, "PaxCount");
                xlWorkSheet.SetCellValue(1, 19, "ProductType");
                xlWorkSheet.SetCellValue(1, 20, "ActionType");
                xlWorkSheet.SetCellValue(1, 21, "BookedBy");
                xlWorkSheet.SetCellValue(1, 22, "Office");
                xlWorkSheet.SetCellValue(1, 23, "ReasonForTravel");
                xlWorkSheet.SetCellValue(1, 24, "RequisitionNumber");
                xlWorkSheet.SetCellValue(1, 25, "OPT");
                xlWorkSheet.SetCellValue(1, 26, "TRID-MarineFare");
                xlWorkSheet.SetCellValue(1, 27, "Verified");
                xlWorkSheet.SetCellValue(1, 28, "Remarks");
                xlWorkSheet.SetCellValue(1, 29, "RegNr");
                xlWorkSheet.SetCellValue(1, 30, "SalesPerson");
                xlWorkSheet.SetCellValue(1, 31, "IssuingAgent");
                xlWorkSheet.SetCellValue(1, 32, "CreatorAgent");

                xlWorkSheet.SetCellStyle(1, 1, 1, 32, mStyles.xlStyleHeader);

                RowCounter = 1;
                var firstRow = mdsDataSet.Tables[0].Rows[0];
                pInvNumber = firstRow[6] == DBNull.Value ? string.Empty : firstRow[6]?.ToString() ?? string.Empty;
                pInvRow = RowCounter + 1;
                pInvSum = 0;
                xlHeaderID = 1;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    RowCounter++;

                    if (Convert.ToString(row[6]) == pInvNumber)
                    {
                        pInvSum += Convert.ToDecimal(row[8]);
                    }
                    else
                    {
                        xlHeaderID++;
                        xlWorkSheet.SetCellValue(pInvRow, 9, pInvSum);
                        for (int k = pInvRow + 1; k < RowCounter; k++)
                        {
                            xlWorkSheet.SetCellStyle(k, 5, k, 9, mStyles.xlStyleGrey);
                            xlWorkSheet.SetCellValue(k, 5, "");
                            xlWorkSheet.SetCellValue(k, 6, "");
                            xlWorkSheet.SetCellValue(k, 7, "");
                            xlWorkSheet.SetCellValue(k, 8, "");
                            xlWorkSheet.SetCellValue(k, 9, "");
                        }
                        pInvSum = Convert.ToDecimal(row[8]);
                        pInvNumber = row[6] == DBNull.Value ? string.Empty : row[6]?.ToString() ?? string.Empty;
                        pInvRow = RowCounter;
                    }

                    xlWorkSheet.SetCellValue(RowCounter, 5, xlHeaderID);

                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToInt32(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDateTime(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDateTime(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDateTime(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToInt32(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[16]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToInt32(row[17]));
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20]));
                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));
                    xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22]));
                    xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23]));
                    xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[24]));
                    xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[25]));
                    xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[26]));
                    xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToString(row[27]));
                    xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToInt32(row[28]));
                    xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToString(row[29]));
                    xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToString(row[30]));
                    xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToString(row[31]));


                }

                if (!string.IsNullOrEmpty(pInvNumber) && pInvSum != 0)
                {
                    xlWorkSheet.SetCellValue(pInvRow, 9, pInvSum);
                    for (int k = pInvRow + 1; k <= RowCounter; k++)
                    {
                        xlWorkSheet.SetCellStyle(k, 5, k, 9, mStyles.xlStyleGrey);
                        xlWorkSheet.SetCellValue(k, 5, "");
                        xlWorkSheet.SetCellValue(k, 6, "");
                        xlWorkSheet.SetCellValue(k, 7, "");
                        xlWorkSheet.SetCellValue(k, 8, "");
                        xlWorkSheet.SetCellValue(k, 9, "");
                    }
                }

                xlWorkSheet.AutoFitColumn(1, 32);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E36_SeaChefs_AllUnits: {ex.Message}", ex);
            }
        }
        public string E38_AirTicketSales()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales");
                xlWorkSheet.FreezePanes(1, 0);



                xlWorkSheet.SetColumnStyle(1, 42, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, 2, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(11, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(35, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(18, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(30, mStyles.xlStyleInteger);

                // Header row
                xlWorkSheet.SetCellValue(1, 1, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 2, "Departure Date");
                xlWorkSheet.SetCellValue(1, 3, "Booked By");
                xlWorkSheet.SetCellValue(1, 4, "Invoice Code");
                xlWorkSheet.SetCellValue(1, 5, "Vessel");
                xlWorkSheet.SetCellValue(1, 6, "Passenger");
                xlWorkSheet.SetCellValue(1, 7, "Routing");
                xlWorkSheet.SetCellValue(1, 8, "Ticketing Airline");
                xlWorkSheet.SetCellValue(1, 9, "Net Payable");
                xlWorkSheet.SetCellValue(1, 10, "Reason For Travel");
                xlWorkSheet.SetCellValue(1, 11, "Issue Date");
                xlWorkSheet.SetCellValue(1, 12, "Client Code");
                xlWorkSheet.SetCellValue(1, 13, "Client Name");
                xlWorkSheet.SetCellValue(1, 14, "Omit");
                xlWorkSheet.SetCellValue(1, 15, "Void");
                xlWorkSheet.SetCellValue(1, 16, "PNR");
                xlWorkSheet.SetCellValue(1, 17, "Ticket Number");
                xlWorkSheet.SetCellValue(1, 18, "Pax Count");
                xlWorkSheet.SetCellValue(1, 19, "Product Type");
                xlWorkSheet.SetCellValue(1, 20, "Action Type");
                xlWorkSheet.SetCellValue(1, 21, "Office/Dept");
                xlWorkSheet.SetCellValue(1, 22, "Cost Centre");
                xlWorkSheet.SetCellValue(1, 23, "Requisition Number");
                xlWorkSheet.SetCellValue(1, 24, "OPT");
                xlWorkSheet.SetCellValue(1, 25, "TRID/MarineFare");
                xlWorkSheet.SetCellValue(1, 26, "Account Code");
                xlWorkSheet.SetCellValue(1, 27, "Verified");
                xlWorkSheet.SetCellValue(1, 28, "Remarks");
                xlWorkSheet.SetCellValue(1, 29, "Transaction Type");
                xlWorkSheet.SetCellValue(1, 30, "RegNr");
                xlWorkSheet.SetCellValue(1, 31, "SalesPerson");
                xlWorkSheet.SetCellValue(1, 32, "Issuing Agent");
                xlWorkSheet.SetCellValue(1, 33, "Creator Agent");
                xlWorkSheet.SetCellValue(1, 34, "Reference");
                xlWorkSheet.SetCellValue(1, 35, "Arrival Date");
                xlWorkSheet.SetCellValue(1, 36, "Connected Document");
                xlWorkSheet.SetCellValue(1, 37, "Pax Remarks");
                xlWorkSheet.SetCellValue(1, 38, "Invoice Status");
                xlWorkSheet.SetCellValue(1, 39, "Other Services");
                xlWorkSheet.SetCellValue(1, 40, "Client Team");
                xlWorkSheet.SetCellValue(1, 41, "Supplier Code");
                xlWorkSheet.SetCellValue(1, 42, "Supplier Name");

                xlWorkSheet.SetCellStyle(1, 1, 1, 42, mStyles.xlStyleHeader);

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    RowCounter++;

                    // Invoice Code
                    string inv = "";
                    if (Convert.ToInt32(row[13]) != 0)
                    {
                        inv = $"{Convert.ToString(row[11])} {Convert.ToString(row[12])} {Convert.ToString(row[13])}".Replace("  ", " ").Trim();
                    }

                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToDateTime(row[35]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[16]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, inv);
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[15]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[30]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[29]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[24]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDateTime(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToInt32(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[17]));

                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[19]));
                    xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[20]));
                    xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[21]));
                    xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[22]));
                    xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[23]));

                    xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[25]));
                    xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToString(row[26]));
                    xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToString(row[27]));
                    xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToInt32(row[28]));

                    xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToString(row[31]));
                    xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToString(row[32]));
                    xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToString(row[33]));
                    xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[34]));

                    xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToDateTime(row[36]));
                    xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToString(row[37]));
                    xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToString(row[38]));

                    if (Convert.ToInt32(row[39]) == 43)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 38, "Cancelled");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 38, mStyles.xlStyleItalic);
                    }
                    else if (!string.IsNullOrEmpty(Convert.ToString(row[40])))
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 38, $"Cancels {Convert.ToString(row[40])}");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 38, mStyles.xlStyleItalic);
                    }

                    xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToString(row[41]));
                    xlWorkSheet.SetCellValue(RowCounter, 40, Convert.ToString(row[42]));
                    xlWorkSheet.SetCellValue(RowCounter, 41, Convert.ToString(row["SuppCode"]));
                    xlWorkSheet.SetCellValue(RowCounter, 42, Convert.ToString(row["SuppName"]));

                    if (!string.IsNullOrEmpty(Convert.ToString(row[3])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 38, mStyles.xlStyleSandyBrown);
                    if (!string.IsNullOrEmpty(Convert.ToString(row[4])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 38, mStyles.xlStyleGrayItalic);
                    if (Convert.ToString(row[10]) == "Refund")
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 38, mStyles.xlStyleRedFont);
                }

                xlWorkSheet.AutoFitColumn(1, 42);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E38_AirTicketSales: {ex.Message}", ex);
            }
        }
        public string E43_DailyProfitReportWithProvisionalAnalysis()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report");
                xlWorkSheet.FreezePanes(3, 2);

                xlWorkSheet.SetColumnStyle(3, 53, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(14, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(27, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(43, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(46, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(48, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(52, mStyles.xlStyleInteger);

                xlWorkSheet.SetColumnStyle(50, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(54, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(1, 3, $"{mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                if (mobjReports.Date2Checked)
                {
                    xlWorkSheet.SetCellValue(1, 47, $"Provisional to {mobjReports.Date2From:dd/MM/yyyy}");
                    xlWorkSheet.SetCellValue(1, 51, "Provisional later");
                }
                else
                {
                    xlWorkSheet.SetCellValue(1, 47, "Provisional");
                }
                xlWorkSheet.SetCellValue(2, 3, "Air Tickets");
                xlWorkSheet.SetCellValue(2, 16, "Other Services");
                xlWorkSheet.SetCellValue(2, 29, "RINVA");
                xlWorkSheet.SetCellValue(2, 32, "Total");

                xlWorkSheet.SetCellValue(3, 1, "Client Code");
                xlWorkSheet.SetCellValue(3, 2, "Client Name");

                for (int i1 = 1; i1 <= 3; i1++)
                {
                    int i = 3;
                    if (i1 == 2) i = 16;
                    if (i1 == 3) i = 32;

                    xlWorkSheet.SetCellValue(3, i, "Net Payable");
                    xlWorkSheet.SetCellValue(3, i + 1, "Net Buy");
                    xlWorkSheet.SetCellValue(3, i + 2, "IW5");
                    xlWorkSheet.SetCellValue(3, i + 3, "IW6");
                    xlWorkSheet.SetCellValue(3, i + 4, "IW7");
                    xlWorkSheet.SetCellValue(3, i + 5, "IW8");
                    xlWorkSheet.SetCellValue(3, i + 6, "IW9");
                    xlWorkSheet.SetCellValue(3, i + 7, "IW11");
                    xlWorkSheet.SetCellValue(3, i + 8, "IW10");
                    xlWorkSheet.SetCellValue(3, i + 9, "IW");
                    xlWorkSheet.SetCellValue(3, i + 10, "Profit");
                    xlWorkSheet.SetCellValue(3, i + 11, "Pax");
                    xlWorkSheet.SetCellValue(3, i + 12, "Profit Per Pax");
                }
                xlWorkSheet.SetCellValue(3, 29, "Net Payable");
                xlWorkSheet.SetCellValue(3, 30, "Net Buy");
                xlWorkSheet.SetCellValue(3, 31, "Profit");

                xlWorkSheet.SetCellValue(2, 45, "YTD");
                xlWorkSheet.SetCellValue(3, 45, "Profit Per Pax");
                xlWorkSheet.SetCellValue(2, 46, "YTD");
                xlWorkSheet.SetCellValue(3, 46, "Pax");

                xlWorkSheet.SetCellValue(2, 47, "Provisional");
                xlWorkSheet.SetCellValue(3, 47, "Uninvoiced T/O");
                var pCommentTO = xlWorkSheet.CreateComment();
                pCommentTO.SetText("Approximate value. Markup or sales fee might still be outstanding");
                xlWorkSheet.InsertComment(2, 47, pCommentTO);

                xlWorkSheet.SetCellValue(3, 48, "Uninvoiced Pax");
                xlWorkSheet.SetCellValue(2, 49, "Provisional Profit");
                xlWorkSheet.SetCellValue(3, 49, "from uninvoiced Pax");
                var pComment = xlWorkSheet.CreateComment();
                pComment.SetText("Approximate value using weighted average");
                xlWorkSheet.InsertComment(2, 49, pComment);
                xlWorkSheet.SetCellValue(3, 50, "To Date");

                if (mobjReports.Date2Checked)
                {
                    xlWorkSheet.SetCellValue(2, 51, "Provisional");
                    xlWorkSheet.SetCellValue(3, 51, "Uninvoiced T/O");
                    xlWorkSheet.SetCellValue(3, 52, "Uninvoiced Pax");
                    xlWorkSheet.SetCellValue(2, 53, "Provisional Profit");
                    xlWorkSheet.SetCellValue(3, 53, "from uninvoiced Pax");
                    xlWorkSheet.SetCellValue(3, 54, "From Date");
                }
                xlWorkSheet.SetCellStyle(1, 3, 3, 54, mStyles.xlStyleHeader);

                xlWorkSheet.MergeWorksheetCells(1, 3, 1, 44);
                xlWorkSheet.MergeWorksheetCells(2, 3, 2, 15);
                xlWorkSheet.MergeWorksheetCells(2, 16, 2, 28);
                xlWorkSheet.MergeWorksheetCells(2, 29, 2, 31);
                xlWorkSheet.MergeWorksheetCells(2, 32, 2, 44);

                xlWorkSheet.MergeWorksheetCells(1, 47, 1, 50);
                xlWorkSheet.MergeWorksheetCells(1, 51, 1, 54);

                RowCounter = 3;
                decimal[] pTotals = new decimal[95];
                decimal[] pOtherClients = new decimal[95];
                int pTopClients = 0;
                int pTopClientsCount = 0;
                decimal pProfitUninvoiced = 0;
                decimal[] pProfitUninvoicedRows = new decimal[mdsDataSet.Tables[0].Rows.Count];
                decimal pProfitUninvoicedLater = 0;
                decimal[] pProfitUninvoicedRowsLater = new decimal[mdsDataSet.Tables[0].Rows.Count];

                // Calculate profit for uninvoiced rows
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    if (row[0].ToString() == "1")
                    {
                        if (!string.IsNullOrEmpty(row[1].ToString()))
                        {
                            pProfitUninvoicedRows[i] = 0;
                            pProfitUninvoicedRowsLater[i] = 0;
                            for (int ii = i + 1; ii < mdsDataSet.Tables[0].Rows.Count; ii++)
                            {
                                var row2 = mdsDataSet.Tables[0].Rows[ii];
                                if (row2[0].ToString() == "2" && row2[1].ToString() == row[1].ToString())
                                {
                                    pProfitUninvoiced = Convert.ToDecimal(row2[81]) * Convert.ToDecimal(row2[82]);
                                    if (pProfitUninvoiced > 0)
                                    {
                                        pProfitUninvoicedRows[i] += pProfitUninvoiced;
                                        pProfitUninvoicedRows[ii] = pProfitUninvoiced;
                                    }
                                    else
                                    {
                                        pProfitUninvoicedRows[ii] = 0;
                                    }
                                    pProfitUninvoicedLater = Convert.ToDecimal(row2[81]) * Convert.ToDecimal(row2[91]);
                                    if (pProfitUninvoicedLater > 0)
                                    {
                                        pProfitUninvoicedRowsLater[i] += pProfitUninvoicedLater;
                                        pProfitUninvoicedRowsLater[ii] = pProfitUninvoicedLater;
                                    }
                                    else
                                    {
                                        pProfitUninvoicedRowsLater[ii] = 0;
                                    }
                                }
                            }
                        }
                        else
                        {
                            pProfitUninvoiced = Convert.ToDecimal(row[81]) * Convert.ToDecimal(row[82]);
                            pProfitUninvoicedRows[i] = pProfitUninvoiced > 0 ? pProfitUninvoiced : 0;
                            pProfitUninvoicedLater = Convert.ToDecimal(row[81]) * Convert.ToDecimal(row[91]);
                            pProfitUninvoicedRowsLater[i] = pProfitUninvoicedLater > 0 ? pProfitUninvoicedLater : 0;
                        }
                    }
                }

                if (int.TryParse(mobjReports.TextEntry, out int topClients))
                    pTopClients = topClients;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    if (row[0].ToString() == "1" && (pTopClients == 0 || pTopClientsCount < pTopClients))
                    {
                        if (mobjReports.BooleanOption1 && !string.IsNullOrEmpty(row[1].ToString()))
                        {
                            int pFirstGroupRow = RowCounter + 1;
                            int pLastGroupRow = RowCounter + 1;
                            for (int ii = i; ii < mdsDataSet.Tables[0].Rows.Count; ii++)
                            {
                                var row2 = mdsDataSet.Tables[0].Rows[ii];
                                if (row2[0].ToString() == "2" && row2[1].ToString() == row[1].ToString())
                                {
                                    RowCounter++;
                                    for (int jj = 2; jj <= 29; jj++)
                                        xlWorkSheet.SetCellValue(RowCounter, jj - 1, Convert.ToDecimal(row2[jj]));
                                    xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDecimal(row2[85]));
                                    xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToDecimal(row2[86]));
                                    xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToDecimal(row2[87]));
                                    for (int jj = 30; jj <= 42; jj++)
                                        xlWorkSheet.SetCellValue(RowCounter, jj + 2, Convert.ToDecimal(row2[jj]));
                                    xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToDecimal(row2[81]));
                                    xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 54, mStyles.xlStyleGrayItalic);
                                    if (Convert.ToDecimal(row2[42]) < Convert.ToDecimal(row2[81]))
                                        xlWorkSheet.SetCellStyle(RowCounter, 44, mStyles.xlStyleRedFontItalic);
                                    xlWorkSheet.SetCellValue(RowCounter, 46, Convert.ToDecimal(row2[80]));
                                    if (Convert.ToDecimal(row2[82]) != 0 || !Convert.IsDBNull(row2[94]))
                                    {
                                        xlWorkSheet.SetCellValue(RowCounter, 47, Convert.ToDecimal(row2[84]));
                                        xlWorkSheet.SetCellValue(RowCounter, 48, Convert.ToDecimal(row2[82]));
                                        if (pProfitUninvoicedRows[ii] != 0)
                                            xlWorkSheet.SetCellValue(RowCounter, 49, pProfitUninvoicedRows[ii]);
                                        xlWorkSheet.SetCellValue(RowCounter, 50, Convert.ToDecimal(row2[94]));
                                    }
                                    if (Convert.ToDecimal(row2[91]) != 0 || !Convert.IsDBNull(row2[95]))
                                    {
                                        xlWorkSheet.SetCellValue(RowCounter, 51, Convert.ToDecimal(row2[93]));
                                        xlWorkSheet.SetCellValue(RowCounter, 52, Convert.ToDecimal(row2[91]));
                                        if (pProfitUninvoicedRowsLater[ii] != 0)
                                            xlWorkSheet.SetCellValue(RowCounter, 53, pProfitUninvoicedRowsLater[ii]);
                                        xlWorkSheet.SetCellValue(RowCounter, 54, Convert.ToDecimal(row2[95]));
                                    }
                                    pLastGroupRow = RowCounter;
                                }
                            }
                            xlWorkSheet.GroupRows(pFirstGroupRow, pLastGroupRow);
                            xlWorkSheet.CollapseRows(pLastGroupRow + 1);
                        }
                        RowCounter++;
                        xlWorkSheet.SetCellValue(RowCounter, 1, (string)row[2]);
                        xlWorkSheet.SetCellValue(RowCounter, 2, (string)row[3]);
                        for (int j = 4; j <= 29; j++)
                        {
                            if (j == 15 || j == 28)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, j - 1, (int)row[j]);
                                pTotals[j] += (int)(row[j]);
                            }
                            else
                            {
                                xlWorkSheet.SetCellValue(RowCounter, j - 1, Convert.ToDecimal(row[j]));
                                pTotals[j] += Convert.ToDecimal(row[j]);
                            }
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDecimal(row[85]));
                        xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToDecimal(row[86]));
                        xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToDecimal(row[87]));
                        pTotals[86] += Convert.ToDecimal(row[85]);
                        pTotals[87] += Convert.ToDecimal(row[86]);
                        pTotals[88] += Convert.ToDecimal(row[87]);
                        for (int j = 30; j <= 42; j++)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, j + 2, Convert.ToDecimal(row[j]));
                            pTotals[j] += Convert.ToDecimal(row[j]);
                        }
                        pTotals[79] += Convert.ToDecimal(row[79]);
                        pTotals[80] += Convert.ToDecimal(row[80]);
                        xlWorkSheet.SetCellValue(RowCounter, 45, Convert.ToDecimal(row[81]));
                        if (Convert.ToDecimal(row[42]) < Convert.ToDecimal(row[81]))
                            xlWorkSheet.SetCellStyle(RowCounter, 44, mStyles.xlStyleRedFont);
                        xlWorkSheet.SetCellValue(RowCounter, 46, Convert.ToDecimal(row[80]));
                        if (Convert.ToDecimal(row[82]) != 0 || !Convert.IsDBNull(row[94]))
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 47, Convert.ToDecimal(row[84]));
                            xlWorkSheet.SetCellValue(RowCounter, 48, Convert.ToDecimal(row[82]));
                            xlWorkSheet.SetCellValue(RowCounter, 49, pProfitUninvoicedRows[i]);
                            if (!Convert.IsDBNull(row[94]))
                                xlWorkSheet.SetCellValue(RowCounter, 50, Convert.ToDateTime(row[94]));
                            pTotals[82] += Convert.ToDecimal(row[82]);
                            pTotals[84] += Convert.ToDecimal(row[84]);
                            pTotals[85] += pProfitUninvoicedRows[i];
                        }
                        if (Convert.ToDecimal(row[91]) != 0 || !Convert.IsDBNull(row[95]))
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 51, Convert.ToDecimal(row[93]));
                            xlWorkSheet.SetCellValue(RowCounter, 52, Convert.ToDecimal(row[91]));
                            xlWorkSheet.SetCellValue(RowCounter, 53, pProfitUninvoicedRowsLater[i]);
                            if (!Convert.IsDBNull(row[95]))
                                xlWorkSheet.SetCellValue(RowCounter, 54, Convert.ToDateTime(row[95]));
                            pTotals[92] += Convert.ToDecimal(row[93]);
                            pTotals[93] += Convert.ToDecimal(row[91]);
                            pTotals[94] += pProfitUninvoicedRowsLater[i];
                        }
                        pTopClientsCount++;
                    }
                    else
                    {
                        for (int j = 2; j <= 42; j++)
                        {
                            if (j > 3)
                                pOtherClients[j] += Convert.ToDecimal(row[j]);
                        }
                        pOtherClients[86] += Convert.ToDecimal(row[85]);
                        pOtherClients[87] += Convert.ToDecimal(row[86]);
                        pOtherClients[88] += Convert.ToDecimal(row[87]);
                        pOtherClients[79] += Convert.ToDecimal(row[79]);
                        pOtherClients[80] += Convert.ToDecimal(row[80]);
                        pOtherClients[82] += Convert.ToDecimal(row[82]);
                        pOtherClients[84] += Convert.ToDecimal(row[84]);
                        pOtherClients[85] += pProfitUninvoicedRows[i];
                        pOtherClients[92] += Convert.ToDecimal(row[93]);
                        pOtherClients[93] += Convert.ToDecimal(row[91]);
                        pOtherClients[94] += pProfitUninvoicedRowsLater[i];
                    }
                }

                // Add "Other Clients" row if needed
                if (pTopClients > 0 && pTopClientsCount >= pTopClients)
                {
                    pOtherClients[16] = pOtherClients[15] != 0 ? pOtherClients[14] / pOtherClients[15] : 0;
                    pOtherClients[29] = pOtherClients[28] != 0 ? pOtherClients[27] / pOtherClients[28] : 0;
                    pOtherClients[42] = pOtherClients[41] != 0 ? pOtherClients[40] / pOtherClients[41] : 0;
                    pOtherClients[81] = pOtherClients[80] != 0 ? pOtherClients[79] / pOtherClients[80] : 0;
                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, "Other Clients");
                    for (int j = 4; j <= 29; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, j - 1, pOtherClients[j]);
                        pTotals[j] += pOtherClients[j];
                    }
                    xlWorkSheet.SetCellValue(RowCounter, 29, pOtherClients[86]);
                    xlWorkSheet.SetCellValue(RowCounter, 30, pOtherClients[87]);
                    xlWorkSheet.SetCellValue(RowCounter, 31, pOtherClients[88]);
                    for (int j = 30; j <= 42; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, j + 2, pOtherClients[j]);
                        pTotals[j] += pOtherClients[j];
                    }
                    pTotals[79] += pOtherClients[79];
                    pTotals[80] += pOtherClients[80];
                    xlWorkSheet.SetCellValue(RowCounter, 45, pOtherClients[81]);
                    if (pOtherClients[42] < pOtherClients[81])
                        xlWorkSheet.SetCellStyle(RowCounter, 44, mStyles.xlStyleRedFont);
                    xlWorkSheet.SetCellValue(RowCounter, 46, pOtherClients[80]);
                    if (pOtherClients[82] != 0 || pOtherClients[84] != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 47, pOtherClients[84]);
                        xlWorkSheet.SetCellValue(RowCounter, 48, pOtherClients[82]);
                        xlWorkSheet.SetCellValue(RowCounter, 49, pOtherClients[85]);
                        pTotals[82] += pOtherClients[82];
                        pTotals[84] += pOtherClients[84];
                        pTotals[85] += pOtherClients[85];
                    }
                    if (pOtherClients[92] != 0 || pOtherClients[93] != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 51, pOtherClients[93]);
                        xlWorkSheet.SetCellValue(RowCounter, 52, pOtherClients[92]);
                        xlWorkSheet.SetCellValue(RowCounter, 53, pOtherClients[94]);
                        pTotals[92] += pOtherClients[92];
                        pTotals[93] += pOtherClients[93];
                        pTotals[94] += pOtherClients[94];
                    }
                }

                // Totals row
                pTotals[16] = pTotals[15] != 0 ? pTotals[14] / pTotals[15] : 0;
                pTotals[29] = pTotals[28] != 0 ? pTotals[27] / pTotals[28] : 0;
                pTotals[42] = pTotals[41] != 0 ? pTotals[40] / pTotals[41] : 0;
                pTotals[81] = pTotals[80] != 0 ? pTotals[79] / pTotals[80] : 0;
                xlWorkSheet.SetCellValue(RowCounter + 1, 2, "TOTAL");
                for (int j = 4; j <= 29; j++)
                    xlWorkSheet.SetCellValue(RowCounter + 1, j - 1, pTotals[j]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 29, pTotals[86]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 30, pTotals[87]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 31, pTotals[88]);
                for (int j = 30; j <= 42; j++)
                    xlWorkSheet.SetCellValue(RowCounter + 1, j + 2, pTotals[j]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 45, pTotals[81]);
                if (pTotals[42] < pTotals[81])
                    xlWorkSheet.SetCellStyle(RowCounter + 1, 44, mStyles.xlStyleRedFont);
                xlWorkSheet.SetCellValue(RowCounter + 1, 46, pTotals[80]);
                xlWorkSheet.SetCellStyle(RowCounter + 1, 2, RowCounter + 1, 54, mStyles.xlStyleHeader);
                if (pTotals[82] != 0 || pTotals[84] != 0)
                {
                    xlWorkSheet.SetCellValue(RowCounter + 1, 47, pTotals[84]);
                    xlWorkSheet.InsertComment(RowCounter + 1, 47, pCommentTO);
                    xlWorkSheet.SetCellValue(RowCounter + 1, 48, pTotals[82]);
                    xlWorkSheet.SetCellValue(RowCounter + 1, 49, pTotals[85]);
                }
                if (pTotals[92] != 0 || pTotals[93] != 0)
                {
                    xlWorkSheet.SetCellValue(RowCounter + 1, 51, pTotals[92]);
                    xlWorkSheet.SetCellValue(RowCounter + 1, 52, pTotals[93]);
                    xlWorkSheet.SetCellValue(RowCounter + 1, 53, pTotals[94]);
                }

                // Final formatting
                mStyles.xlStyleHeader.Fill.SetPatternForegroundColor(System.Drawing.Color.LightGray);
                mStyles.xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General);
                xlWorkSheet.SetCellStyle(3, 51, RowCounter, 54, mStyles.xlStyleHeader);

                xlWorkSheet.SetCellStyle(3, 15, RowCounter + 1, 15, mStyles.xlStyleDecimalYellow);
                xlWorkSheet.SetCellStyle(3, 28, RowCounter + 1, 28, mStyles.xlStyleDecimalYellow);
                xlWorkSheet.SetCellStyle(3, 29, RowCounter + 1, 31, mStyles.xlStyleLightGreen);
                xlWorkSheet.SetCellStyle(3, 44, RowCounter + 1, 45, mStyles.xlStyleDecimalYellow);

                xlWorkSheet.SetColumnStyle(3, 53, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(14, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(27, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(43, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(46, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(48, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(52, mStyles.xlStyleInteger);

                xlWorkSheet.SetColumnStyle(50, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(54, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(5, 11, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetColumnStyle(18, 24, mStyles.xlStyleGrayItalic);
                xlWorkSheet.SetColumnStyle(34, 40, mStyles.xlStyleGrayItalic);

                xlWorkSheet.AutoFitColumn(0, 54);
                xlWorkSheet.GroupColumns(5, 11);
                xlWorkSheet.CollapseColumns(12);
                xlWorkSheet.GroupColumns(18, 24);
                xlWorkSheet.CollapseColumns(25);
                xlWorkSheet.GroupColumns(34, 40);
                xlWorkSheet.CollapseColumns(41);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string E45_AirTicketSalesAll()
        {

            RowCounter = 1;

            try
            {
                xlWorkSheet.SetColumnStyle(1, 133, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(17, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(36, 38, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(133, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(15, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(28, 29, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(39, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(132, mStyles.xlStyleTime);

                // Header row
                string[] headers = {
            "Issue Date", "Client Code", "Client Name", "Omit", "Void", "PNR", "Ticket Number", "Passenger", "Pax Count", "Product Type", "Action Type",
            "Inv Code", "Inv Series", "Inv Number", "Invoice Date", "Vessel", "Net Payable", "Verified", "Remarks", "Transaction Type", "RegNr",
            "Ticketing Airline", "Routing", "SalesPerson", "Issuing Agent", "Creator Agent", "Reference", "Departure Date", "Arrival Date", "Connected Document",
            "Pax Remarks", "Doc Status ID", "Cancels Docs", "Sertvices Description", "Client Team", "Sell", "Buy", "Profit", "PaxCount+-",
            // 40-133: Custom fields, labels, mandatory, values, etc.
            "01-Label", "01-Mandatory", "01-BookedBy (G)", "02-Label", "02-Mandatory", "02-Office", "04-Label", "04-Mandatory", "04-ReasonForTravel",
            "05-Label", "05-Mandatory", "05-CostCentre", "06-Label", "06-Mandatory", "06-Savings (G)", "07-Label", "07-Mandatory", "07-Losses (G)",
            "08-Label", "08-Mandatory", "08-Savings/Losses Reason (G)", "09-Label", "09-Mandatory", "09-Travel Definition (G)", "11-Label", "11-Mandatory", "11-RequisitionNumber",
            "12-Label", "12-Mandatory", "12-Passenger ID (G)", "13-Label", "13-Mandatory", "13-OPT (G)", "14-Label", "14-Mandatory", "14-TRID-MarineFare",
            "15-Label", "15-Mandatory", "15-AccountCode", "16-Label", "16-Mandatory", "16-Rank", "17-Label", "17-Mandatory", "17-Nationality",
            "18-Label", "18-Mandatory", "18-REF1", "19-Label", "19-Mandatory", "19-REF2", "20-Label", "20-Mandatory", "20-REF3",
            "21-Label", "21-Mandatory", "21-REF4", "22-Label", "22-Mandatory", "22-REF5", "23-Label", "23-Mandatory", "23-REF6",
            "24-Label", "24-Mandatory", "24-REF7", "25-Label", "25-Mandatory", "25-REF8", "26-Label", "26-Mandatory", "26-REF9",
            "27-Label", "27-Mandatory", "27-REF10", "28-Label", "28-Mandatory", "28-REF11", "29-Label", "29-Mandatory", "29-REF12",
            "30-Label", "30-Mandatory", "30-REF13", "31-Label", "31-Mandatory", "31-REF14", "32-Label", "32-Mandatory", "32-REF15",
            "Fare Basis", "Cabin Class", "Duration", "CO2"
        };
                for (int h = 0; h < headers.Length; h++)
                    xlWorkSheet.SetCellValue(1, h + 1, headers[h]);
                xlWorkSheet.SetCellStyle(1, 1, 1, 133, mStyles.xlStyleHeader);

                // Data rows
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    RowCounter++;

                    // Main fields (0-10)
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToInt32(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10]));



                    // Inv fields (11-14) only if InvNumber != 0
                    if (Convert.ToInt32(row[13]) != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11]));
                        xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[12]));
                        xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToInt32(row[13]));
                        xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDateTime(row[14]));
                    }

                    // 15-31
                    for (int j = 15; j <= 31; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[16]));
                        xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17]));
                        xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToInt32(row[20]));
                        xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21]));
                        xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22]));
                        xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23]));
                        xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[24]));
                        xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[25]));
                        xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[26]));
                        xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToDateTime(row[27]));
                        xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDateTime(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToString(row[29]));
                        xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToString(row[30]));
                        xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToInt32(row[31]));

                    }

                    // Cancelled/Cancelled Docs
                    if (Convert.ToInt32(row[31]) == 43)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 33, "Cancelled");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 129, mStyles.xlStyleItalic);
                    }
                    else if (!string.IsNullOrEmpty(Convert.ToString(row[32])))
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 33, $"Cancels {Convert.ToString(row[32])}");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 129, mStyles.xlStyleItalic);
                    }

                    // 33-38
                    for (int j = 33; j <= 38; j++)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[33]));
                        xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToString(row[34]));
                        xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToDecimal(row[35]));
                        xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToDecimal(row[36]));
                        xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToDecimal(row[37]));
                        xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToInt32(row[38]));
                    }

                    // Custom fields (labels, mandatory, value) 39-126 step 3
                    for (int j = 39; j <= 126; j += 3)
                    {
                        bool isEmpty = string.IsNullOrEmpty(Convert.ToString(row[j])) && string.IsNullOrEmpty(Convert.ToString(row[j + 2]));
                        if (isEmpty)
                        {
                            xlWorkSheet.SetCellStyle(RowCounter, j + 1, RowCounter, j + 3, mStyles.xlStyleLightGray);
                        }
                        else
                        {
                            if (Convert.ToString(row[j + 1]) == "Mandatory")
                                xlWorkSheet.SetCellStyle(RowCounter, j + 1, RowCounter, j + 3, mStyles.xlStyleYellow);

                            xlWorkSheet.SetCellValue(RowCounter, j + 1, Convert.ToString(row[j]));
                            xlWorkSheet.SetCellValue(RowCounter, j + 2, Convert.ToString(row[j + 1]));
                            xlWorkSheet.SetCellValue(RowCounter, j + 3, Convert.ToString(row[j + 2]));
                        }
                    }

                    // Remaining fields
                    xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[33]));
                    xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToString(row[34]));
                    xlWorkSheet.SetCellValue(RowCounter, 130, Convert.ToString(row[129]));
                    xlWorkSheet.SetCellValue(RowCounter, 131, Convert.ToString(row[130]));
                    xlWorkSheet.SetCellValue(RowCounter, 132, row[131].ToString());
                    xlWorkSheet.SetCellValue(RowCounter, 133, Convert.ToString(row[132]));

                    // Style for Omit/Void/Refund
                    if (!string.IsNullOrEmpty(Convert.ToString(row[3])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 130, mStyles.xlStyleSandyBrown);
                    if (!string.IsNullOrEmpty(Convert.ToString(row[4])))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 130, mStyles.xlStyleGrayItalic);
                    if (Convert.ToString(row[10]) == "Refund")
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 130, mStyles.xlStyleRedFont);
                }

                xlWorkSheet.AutoFitColumn(1, 133);
                for (int ic = 40; ic <= 127; ic += 3)
                {
                    xlWorkSheet.GroupColumns(ic, ic + 1);
                    xlWorkSheet.CollapseColumns(ic + 2);
                }

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E45_AirTicketSalesAll: " + ex.Message, ex);
            }
        }
        public string E47_DailyProfitReportTotalsOnly()
        {

            RowCounter = 3;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Totals");
                xlWorkSheet.FreezePanes(3, 2);



                xlWorkSheet.SetColumnStyle(3, 6, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(10, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(7, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);

                xlWorkSheet.SetCellValue(1, 3, $"{mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                xlWorkSheet.SetCellValue(2, 3, "Total");
                xlWorkSheet.SetCellValue(3, 1, "Client Code");
                xlWorkSheet.SetCellValue(3, 2, "Client Name");
                xlWorkSheet.SetCellValue(3, 3, "Net Payable");
                xlWorkSheet.SetCellValue(3, 4, "Net Buy");
                xlWorkSheet.SetCellValue(3, 5, "IW");
                xlWorkSheet.SetCellValue(3, 6, "Profit");
                xlWorkSheet.SetCellValue(3, 7, "Total Pax");
                xlWorkSheet.SetCellValue(3, 8, "IW5");
                xlWorkSheet.SetCellValue(3, 9, "IW5 Pax");
                xlWorkSheet.SetCellValue(3, 10, "IW5/Pax");

                xlWorkSheet.SetCellStyle(1, 1, 3, 10, mStyles.xlStyleHeader);
                xlWorkSheet.MergeWorksheetCells(1, 3, 1, 10);
                xlWorkSheet.MergeWorksheetCells(2, 3, 2, 10);

                if (mobjReports.BooleanOption1 == false)
                {
                    xlWorkSheet.SetCellValue(1, 1, "INCLUDING OMIT");
                    xlWorkSheet.SetCellStyle(1, 1, mStyles.xlStyleSandyBrown);
                }

                decimal[] pTotals = new decimal[101];

                for (int ii = 0; ii < mdsDataSet.Tables[0].Rows.Count; ii++)
                {
                    var row = mdsDataSet.Tables[0].Rows[ii];

                    if (Convert.ToString(row[2]) != "")
                    {
                        if (Convert.ToDecimal(row[30]) != 0 ||
                            Convert.ToDecimal(row[31]) != 0 ||
                            Convert.ToDecimal(row[32]) != 0 ||
                            Convert.ToDecimal(row[39]) != 0 ||
                            Convert.ToDecimal(row[40]) != 0 ||
                            Convert.ToDecimal(row[41]) != 0 ||
                            Convert.ToDecimal(row[100]) != 0)
                        {
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[2]));
                            xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[3]));
                            xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToDecimal(row[30]));
                            xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDecimal(row[31]));
                            xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDecimal(row[39]));
                            xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row[40]));
                            xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[41]));
                            xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[32]));
                            xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[100]));
                            if (Convert.ToDecimal(row[100]) != 0)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[32]) / Convert.ToDecimal(row[100]));
                                if (Convert.ToDecimal(row[39]) != Convert.ToDecimal(row[32]))
                                {
                                    xlWorkSheet.SetCellValue(RowCounter, 11, $"IW5 difference {(Convert.ToDecimal(row[39]) - Convert.ToDecimal(row[32])):#,##0.00}");
                                }
                            }
                            pTotals[30] += Convert.ToDecimal(row[30]);
                            pTotals[31] += Convert.ToDecimal(row[31]);
                            pTotals[32] += Convert.ToDecimal(row[32]);
                            pTotals[39] += Convert.ToDecimal(row[39]);
                            pTotals[40] += Convert.ToDecimal(row[40]);
                            pTotals[41] += Convert.ToDecimal(row[41]);
                            pTotals[100] += Convert.ToDecimal(row[100]);
                        }
                    }
                }

                xlWorkSheet.Sort(4, 1, RowCounter, 11, 6, false);
                xlWorkSheet.SetCellStyle(4, 8, RowCounter, 10, mStyles.xlStyleGrey);

                xlWorkSheet.SetCellValue(RowCounter + 1, 2, "TOTAL");
                xlWorkSheet.SetCellValue(RowCounter + 1, 3, pTotals[30]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 4, pTotals[31]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 5, pTotals[39]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 6, pTotals[40]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 7, pTotals[41]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 8, pTotals[32]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 9, pTotals[100]);
                if (pTotals[100] != 0)
                {
                    xlWorkSheet.SetCellValue(RowCounter + 1, 10, pTotals[32] / pTotals[100]);
                }
                xlWorkSheet.SetCellStyle(RowCounter + 1, 2, RowCounter + 1, 10, mStyles.xlStyleHeader);

                xlWorkSheet.AutoFitColumn(0, 11);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E47_DailyProfitReportTotalsOnly: " + ex.Message, ex);
            }
        }
        public string E48_DailyProfitReportTotalsPerInvoice()
        {

            RowCounter = 3;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Totals");
                xlWorkSheet.FreezePanes(3, 2);

                xlWorkSheet.SetColumnStyle(7, 10, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(12, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(14, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(11, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(13, mStyles.xlStyleInteger);

                xlWorkSheet.SetColumnStyle(5, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(1, 3, $"{mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                xlWorkSheet.SetCellValue(2, 3, "Total");
                xlWorkSheet.SetCellValue(3, 1, "Client Code");
                xlWorkSheet.SetCellValue(3, 2, "Client Name");
                xlWorkSheet.SetCellValue(3, 3, "Doc Type");
                xlWorkSheet.SetCellValue(3, 4, "Doc Number");
                xlWorkSheet.SetCellValue(3, 5, "Doc Date");
                xlWorkSheet.SetCellValue(3, 6, "Omit");
                xlWorkSheet.SetCellValue(3, 7, "Net Payable");
                xlWorkSheet.SetCellValue(3, 8, "Net Buy");
                xlWorkSheet.SetCellValue(3, 9, "IW");
                xlWorkSheet.SetCellValue(3, 10, "Profit");
                xlWorkSheet.SetCellValue(3, 11, "Total Pax");
                xlWorkSheet.SetCellValue(3, 12, "IW5");
                xlWorkSheet.SetCellValue(3, 13, "IW5 Pax");
                xlWorkSheet.SetCellValue(3, 14, "IW5/Pax");
                xlWorkSheet.SetCellValue(3, 15, "Notes");
                xlWorkSheet.SetCellStyle(1, 1, 3, 15, mStyles.xlStyleHeader);
                xlWorkSheet.MergeWorksheetCells(1, 3, 1, 15);
                xlWorkSheet.MergeWorksheetCells(2, 3, 2, 15);

                if (!mobjReports.BooleanOption1)
                {
                    xlWorkSheet.SetCellValue(1, 1, "INCLUDING OMIT");
                    xlWorkSheet.SetCellStyle(1, 1, mStyles.xlStyleSandyBrown);
                }

                decimal[] pTotals = new decimal[100];

                for (int ii = 0; ii < mdsDataSet.Tables[0].Rows.Count; ii++)
                {
                    var row = mdsDataSet.Tables[0].Rows[ii];
                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDateTime(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToInt32(row[7]));

                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[34]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[35]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[43]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[44]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[45]));
                    if (Convert.ToDecimal(row[36]) != 0)
                        xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[36]));
                    if (Convert.ToDecimal(row[52]) != 0)
                        xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[52]));
                    if (Convert.ToDecimal(row[52]) != 0)
                        xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDecimal(row[36]) / Convert.ToDecimal(row[52]));
                    if (Convert.ToDecimal(row[43]) != Convert.ToDecimal(row[36]))
                        xlWorkSheet.SetCellValue(RowCounter, 15, $"IW5 difference {Convert.ToDecimal(row[43]) - Convert.ToDecimal(row[36]):#,##0.00}");

                    pTotals[34] += Convert.ToDecimal(row[34]);
                    pTotals[35] += Convert.ToDecimal(row[35]);
                    pTotals[43] += Convert.ToDecimal(row[43]);
                    pTotals[44] += Convert.ToDecimal(row[44]);
                    pTotals[45] += Convert.ToDecimal(row[45]);
                    pTotals[36] += Convert.ToDecimal(row[36]);
                    pTotals[52] += Convert.ToDecimal(row[52]);

                    if (Convert.ToInt32(row[7]) == 1)
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 14, mStyles.xlStyleSandyBrown);
                }

                xlWorkSheet.Sort(4, 1, RowCounter, 15, 4, true);
                xlWorkSheet.Sort(4, 1, RowCounter, 15, 3, true);
                xlWorkSheet.Sort(4, 1, RowCounter, 15, 1, true);
                xlWorkSheet.SetCellStyle(4, 12, RowCounter, 15, mStyles.xlStyleGrey);

                xlWorkSheet.SetCellValue(RowCounter + 1, 2, "TOTAL");
                xlWorkSheet.SetCellValue(RowCounter + 1, 7, pTotals[34]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 8, pTotals[35]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 9, pTotals[43]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 10, pTotals[44]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 11, pTotals[45]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 12, pTotals[36]);
                xlWorkSheet.SetCellValue(RowCounter + 1, 13, pTotals[52]);
                if (pTotals[52] != 0)
                    xlWorkSheet.SetCellValue(RowCounter + 1, 14, pTotals[36] / pTotals[52]);
                xlWorkSheet.SetCellStyle(RowCounter + 1, 2, RowCounter + 1, 14, mStyles.xlStyleHeader);

                xlWorkSheet.AutoFitColumn(0, 15);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E48_DailyProfitReportTotalsPerInvoice: " + ex.Message, ex);
            }
        }
        public string E51_Daily_Profit_Totals_per_Category()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report");
                xlWorkSheet.FreezePanes(5, 4);



                if (!mobjReports.BooleanOption1)
                {
                    xlWorkSheet.SetCellValue(1, 2, "INCLUDING OMIT");
                    xlWorkSheet.SetCellStyle(1, 2, mStyles.xlStyleSandyBrown);
                }
                xlWorkSheet.SetCellValue(2, 2, $"{mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                xlWorkSheet.SetCellValue(1, 5, "Total");
                xlWorkSheet.MergeWorksheetCells(1, 5, 1, 8);
                xlWorkSheet.SetCellValue(1, 9, "Total Air");
                xlWorkSheet.MergeWorksheetCells(1, 9, 1, 12);
                xlWorkSheet.SetCellValue(1, 13, "Total Services");
                xlWorkSheet.MergeWorksheetCells(1, 13, 1, 16);
                xlWorkSheet.SetCellValue(1, 17, "AIR");
                xlWorkSheet.MergeWorksheetCells(1, 17, 1, 36);
                xlWorkSheet.SetCellValue(1, 37, "SERVICES");
                xlWorkSheet.MergeWorksheetCells(1, 37, 1, 59);

                for (int ii = 0; ii <= 1; ii++)
                {
                    xlWorkSheet.SetCellValue(2, ii * 20 + 17, "01 Marine");
                    xlWorkSheet.MergeWorksheetCells(2, ii * 20 + 17, 2, ii * 20 + 20);
                    xlWorkSheet.SetCellValue(2, ii * 20 + 21, "02 Interoffice");
                    xlWorkSheet.MergeWorksheetCells(2, ii * 20 + 21, 2, ii * 20 + 24);
                    xlWorkSheet.SetCellValue(2, ii * 20 + 25, "03 Corporate");
                    xlWorkSheet.MergeWorksheetCells(2, ii * 20 + 25, 2, ii * 20 + 28);
                    xlWorkSheet.SetCellValue(2, ii * 20 + 29, "04 Non Marine");
                    xlWorkSheet.MergeWorksheetCells(2, ii * 20 + 29, 2, ii * 20 + 32);
                    xlWorkSheet.SetCellValue(2, ii * 20 + 33, "05 Care of");
                    xlWorkSheet.MergeWorksheetCells(2, ii * 20 + 33, 2, ii * 20 + 36);
                }
                xlWorkSheet.SetCellValue(2, 57, "RINVA");
                xlWorkSheet.MergeWorksheetCells(2, 57, 2, 59);

                xlWorkSheet.SetCellValue(5, 1, "Tots");
                xlWorkSheet.SetCellValue(5, 2, "Client Group");
                xlWorkSheet.SetCellValue(5, 3, "Client Code");
                xlWorkSheet.SetCellValue(5, 4, "Client Name");
                for (int ii = 0; ii <= 12; ii++)
                {
                    xlWorkSheet.SetCellValue(5, ii * 4 + 5, "Net Payable");
                    xlWorkSheet.SetCellValue(5, ii * 4 + 6, "Net Buy");
                    xlWorkSheet.SetCellValue(5, ii * 4 + 7, "Profit");
                    xlWorkSheet.SetCellValue(5, ii * 4 + 8, "Pax");
                    xlWorkSheet.SetColumnStyle(ii * 4 + 5, ii * 4 + 7, mStyles.xlStyleDecimal);
                    xlWorkSheet.SetColumnStyle(ii * 4 + 8, mStyles.xlStyleInteger);
                }
                xlWorkSheet.SetCellValue(5, 57, "Net Payable");
                xlWorkSheet.SetCellValue(5, 58, "Net Buy");
                xlWorkSheet.SetCellValue(5, 59, "Profit");
                xlWorkSheet.SetColumnStyle(57, 59, mStyles.xlStyleDecimal);

                xlWorkSheet.SetCellStyle(1, 1, 5, 59, mStyles.xlStyleHeader);

                RowCounter = 5;
                decimal[] Totals = new decimal[60];
                decimal[] TotClientGroup = new decimal[60];
                int TotsLevel;
                string PrevClientGroup = "";
                int FirstGroupRow = 0;
                int LastGroupRow = 0;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    int Category = 0;
                    switch (Convert.ToString(row[2]))
                    {
                        case "01": Category = 1; break;
                        case "02": Category = 2; break;
                        case "03": Category = 3; break;
                        case "04": Category = 4; break;
                        default: Category = 5; break;
                    }
                    TotsLevel = Convert.ToInt32(row[0]);

                    if (TotsLevel == 2 || (TotsLevel == 1 && Convert.ToString(row[3]) != ""))
                    {
                        if (TotsLevel == 2 && PrevClientGroup != Convert.ToString(row[1]))
                        {
                            if (!string.IsNullOrEmpty(PrevClientGroup))
                            {
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 2, PrevClientGroup);
                                for (int j = 5; j <= 59; j++)
                                {
                                    xlWorkSheet.SetCellValue(RowCounter, j, TotClientGroup[j]);
                                    TotClientGroup[j] = 0;
                                }
                            }
                            xlWorkSheet.GroupRows(FirstGroupRow, LastGroupRow);
                            xlWorkSheet.CollapseRows(LastGroupRow + 1);
                            FirstGroupRow = RowCounter + 1;
                        }
                        RowCounter++;
                        LastGroupRow = RowCounter;
                        if (TotsLevel == 2)
                        {
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 59, mStyles.xlStyleGrayItalic);
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToInt32(row[0]));
                        xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1]));
                        xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[3]));
                        xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[4]));
                        // Total
                        xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDecimal(row[5]) + Convert.ToDecimal(row[18]) + Convert.ToDecimal(row[86]));
                        xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row[6]) + Convert.ToDecimal(row[19]) + Convert.ToDecimal(row[87]));
                        xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToDecimal(row[15]) + Convert.ToDecimal(row[28]) + Convert.ToDecimal(row[88]));
                        xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDecimal(row[16]) + Convert.ToDecimal(row[29]));

                        // AIR
                        xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[5]));
                        xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[6]));
                        xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[16]));

                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 13, Convert.ToDecimal(row[5]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 14, Convert.ToDecimal(row[6]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 15, Convert.ToDecimal(row[15]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 16, Convert.ToDecimal(row[16]));
                        // Services
                        xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDecimal(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDecimal(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[29]));

                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 33, Convert.ToDecimal(row[18]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 34, Convert.ToDecimal(row[19]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 35, Convert.ToDecimal(row[28]));
                        xlWorkSheet.SetCellValue(RowCounter, Category * 4 + 36, Convert.ToDecimal(row[29]));

                        // RINVA
                        xlWorkSheet.SetCellValue(RowCounter, 57, Convert.ToDecimal(row[86]));
                        xlWorkSheet.SetCellValue(RowCounter, 58, Convert.ToDecimal(row[87]));
                        xlWorkSheet.SetCellValue(RowCounter, 59, Convert.ToDecimal(row[88]));

                        // TOTALS
                        Totals[5] += Convert.ToDecimal(row[5]) + Convert.ToDecimal(row[18]) + Convert.ToDecimal(row[86]);
                        Totals[6] += Convert.ToDecimal(row[6]) + Convert.ToDecimal(row[19]) + Convert.ToDecimal(row[87]);
                        Totals[7] += Convert.ToDecimal(row[15]) + Convert.ToDecimal(row[28]) + Convert.ToDecimal(row[88]);
                        Totals[8] += Convert.ToDecimal(row[16]) + Convert.ToDecimal(row[29]);
                        // AIR
                        Totals[9] += Convert.ToDecimal(row[5]);
                        Totals[10] += Convert.ToDecimal(row[6]);
                        Totals[11] += Convert.ToDecimal(row[15]);
                        Totals[12] += Convert.ToDecimal(row[16]);
                        Totals[Category * 4 + 13] += Convert.ToDecimal(row[5]);
                        Totals[Category * 4 + 14] += Convert.ToDecimal(row[6]);
                        Totals[Category * 4 + 15] += Convert.ToDecimal(row[15]);
                        Totals[Category * 4 + 16] += Convert.ToDecimal(row[16]);
                        // SERVICES
                        Totals[13] += Convert.ToDecimal(row[18]);
                        Totals[14] += Convert.ToDecimal(row[19]);
                        Totals[15] += Convert.ToDecimal(row[28]);
                        Totals[16] += Convert.ToDecimal(row[29]);
                        Totals[Category * 4 + 33] += Convert.ToDecimal(row[18]);
                        Totals[Category * 4 + 34] += Convert.ToDecimal(row[19]);
                        Totals[Category * 4 + 35] += Convert.ToDecimal(row[28]);
                        Totals[Category * 4 + 36] += Convert.ToDecimal(row[29]);
                        // RINVA
                        Totals[57] += Convert.ToDecimal(row[86]);
                        Totals[58] += Convert.ToDecimal(row[87]);
                        Totals[59] += Convert.ToDecimal(row[88]);

                        if (TotsLevel == 2)
                        {
                            // TOTALS
                            TotClientGroup[5] += Convert.ToDecimal(row[5]) + Convert.ToDecimal(row[18]) + Convert.ToDecimal(row[86]);
                            TotClientGroup[6] += Convert.ToDecimal(row[6]) + Convert.ToDecimal(row[19]) + Convert.ToDecimal(row[87]);
                            TotClientGroup[7] += Convert.ToDecimal(row[15]) + Convert.ToDecimal(row[28]) + Convert.ToDecimal(row[88]);
                            TotClientGroup[8] += Convert.ToDecimal(row[16]) + Convert.ToDecimal(row[29]);
                            // AIR
                            TotClientGroup[9] += Convert.ToDecimal(row[5]);
                            TotClientGroup[10] += Convert.ToDecimal(row[6]);
                            TotClientGroup[11] += Convert.ToDecimal(row[15]);
                            TotClientGroup[12] += Convert.ToDecimal(row[16]);
                            TotClientGroup[Category * 4 + 13] += Convert.ToDecimal(row[5]);
                            TotClientGroup[Category * 4 + 14] += Convert.ToDecimal(row[6]);
                            TotClientGroup[Category * 4 + 15] += Convert.ToDecimal(row[15]);
                            TotClientGroup[Category * 4 + 16] += Convert.ToDecimal(row[16]);
                            // SERVICES
                            TotClientGroup[13] += Convert.ToDecimal(row[18]);
                            TotClientGroup[14] += Convert.ToDecimal(row[19]);
                            TotClientGroup[15] += Convert.ToDecimal(row[28]);
                            TotClientGroup[16] += Convert.ToDecimal(row[29]);
                            TotClientGroup[Category * 4 + 33] += Convert.ToDecimal(row[18]);
                            TotClientGroup[Category * 4 + 34] += Convert.ToDecimal(row[19]);
                            TotClientGroup[Category * 4 + 35] += Convert.ToDecimal(row[28]);
                            TotClientGroup[Category * 4 + 36] += Convert.ToDecimal(row[29]);
                            // RINVA
                            TotClientGroup[57] += Convert.ToDecimal(row[86]);
                            TotClientGroup[58] += Convert.ToDecimal(row[87]);
                            TotClientGroup[59] += Convert.ToDecimal(row[88]);

                            PrevClientGroup = row[1] == DBNull.Value ? string.Empty : row[1]?.ToString() ?? string.Empty;
                        }
                    }
                }
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 2, PrevClientGroup);
                for (int j = 5; j <= 59; j++)
                {
                    xlWorkSheet.SetCellValue(RowCounter, j, TotClientGroup[j]);
                    TotClientGroup[j] = 0;
                }
                xlWorkSheet.GroupRows(FirstGroupRow, LastGroupRow);
                xlWorkSheet.CollapseRows(LastGroupRow + 1);

                for (int i = 5; i <= 59; i++)
                {
                    xlWorkSheet.SetCellValue(3, i, Totals[i]);
                }

                xlWorkSheet.SetCellStyle(1, 5, RowCounter, 8, mStyles.xlStyleLightSteelBlue);
                xlWorkSheet.SetCellStyle(1, 9, RowCounter, 12, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(1, 13, RowCounter, 16, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(1, 17, RowCounter, 20, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(1, 21, RowCounter, 24, mStyles.xlStyleLightGreen);
                xlWorkSheet.SetCellStyle(1, 25, RowCounter, 28, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(1, 29, RowCounter, 32, mStyles.xlStyleLightGreen);
                xlWorkSheet.SetCellStyle(1, 33, RowCounter, 36, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(1, 37, RowCounter, 40, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(1, 41, RowCounter, 44, mStyles.xlStyleKhaki);
                xlWorkSheet.SetCellStyle(1, 45, RowCounter, 48, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(1, 49, RowCounter, 52, mStyles.xlStyleKhaki);
                xlWorkSheet.SetCellStyle(1, 53, RowCounter, 56, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(1, 57, RowCounter, 59, mStyles.xlStyleKhaki);

                xlWorkSheet.DrawBorderGrid(1, 5, RowCounter, 59, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.DrawBorderGrid(5, 1, RowCounter, 4, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.AutoFitColumn(1, 59);

                xlWorkSheet.AddWorksheet("Summary");
                xlWorkSheet.SetCellValue(1, 3, "Net Payable");
                xlWorkSheet.SetCellValue(1, 4, "Net Buy");
                xlWorkSheet.SetCellValue(1, 5, "Profit");
                xlWorkSheet.SetCellValue(1, 6, "Pax");

                xlWorkSheet.SetCellValue(2, 1, "Total");

                xlWorkSheet.SetCellValue(4, 1, "AIR");
                xlWorkSheet.MergeWorksheetCells(4, 1, 9, 1);
                xlWorkSheet.SetCellValue(11, 1, "SERVICES");
                xlWorkSheet.MergeWorksheetCells(11, 1, 16, 1);

                for (int i = 4; i <= 11; i += 7)
                {
                    xlWorkSheet.SetCellValue(i, 2, "Total");
                    xlWorkSheet.SetCellValue(i + 1, 2, "01 Marine");
                    xlWorkSheet.SetCellValue(i + 2, 2, "02 Interoffice");
                    xlWorkSheet.SetCellValue(i + 3, 2, "03 Corporate");
                    xlWorkSheet.SetCellValue(i + 4, 2, "04 Non marine");
                    xlWorkSheet.SetCellValue(i + 5, 2, "05 Care of");
                }
                xlWorkSheet.SetCellValue(18, 2, "RINVA");

                xlWorkSheet.SetCellValue(2, 3, Totals[5]);
                xlWorkSheet.SetCellValue(2, 4, Totals[6]);
                xlWorkSheet.SetCellValue(2, 5, Totals[7]);
                xlWorkSheet.SetCellValue(2, 6, Totals[8]);

                xlWorkSheet.SetCellValue(4, 3, Totals[9]);
                xlWorkSheet.SetCellValue(4, 4, Totals[10]);
                xlWorkSheet.SetCellValue(4, 5, Totals[11]);
                xlWorkSheet.SetCellValue(4, 6, Totals[12]);

                xlWorkSheet.SetCellValue(11, 3, Totals[13]);
                xlWorkSheet.SetCellValue(11, 4, Totals[14]);
                xlWorkSheet.SetCellValue(11, 5, Totals[15]);
                xlWorkSheet.SetCellValue(11, 6, Totals[16]);

                for (int j = 0; j <= 1; j++)
                {
                    for (int i = 0; i <= 4; i++)
                    {
                        xlWorkSheet.SetCellValue(j * 7 + i + 5, 3, Totals[j * 20 + i * 4 + 17]);
                        xlWorkSheet.SetCellValue(j * 7 + i + 5, 4, Totals[j * 20 + i * 4 + 18]);
                        xlWorkSheet.SetCellValue(j * 7 + i + 5, 5, Totals[j * 20 + i * 4 + 19]);
                        xlWorkSheet.SetCellValue(j * 7 + i + 5, 6, Totals[j * 20 + i * 4 + 20]);
                    }
                }

                xlWorkSheet.SetCellValue(18, 3, Totals[57]);
                xlWorkSheet.SetCellValue(18, 4, Totals[58]);
                xlWorkSheet.SetCellValue(18, 5, Totals[59]);

                xlWorkSheet.SetColumnStyle(3, 5, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(6, mStyles.xlStyleInteger);

                mStyles.xlStyleLightSteelBlue.Font.Bold = true;
                xlWorkSheet.SetCellStyle(1, 1, 2, 6, mStyles.xlStyleLightSteelBlue);
                xlWorkSheet.SetCellStyle(4, 2, 4, 6, mStyles.xlStyleLightSteelBlue);
                xlWorkSheet.SetCellStyle(11, 2, 11, 6, mStyles.xlStyleLightSteelBlue);

                xlWorkSheet.SetCellStyle(5, 2, 5, 6, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(6, 2, 6, 6, mStyles.xlStyleLightGreen);
                xlWorkSheet.SetCellStyle(7, 2, 7, 6, mStyles.xlStyleHoneyDew);
                xlWorkSheet.SetCellStyle(8, 2, 8, 6, mStyles.xlStyleLightGreen);
                xlWorkSheet.SetCellStyle(9, 2, 9, 6, mStyles.xlStyleHoneyDew);
                mStyles.xlStyleHoneyDew.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
                xlWorkSheet.SetCellStyle(4, 1, 4, 1, mStyles.xlStyleHoneyDew);

                xlWorkSheet.SetCellStyle(12, 2, 12, 6, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(13, 2, 13, 6, mStyles.xlStyleKhaki);
                xlWorkSheet.SetCellStyle(14, 2, 14, 6, mStyles.xlStyleLemonChiffon);
                xlWorkSheet.SetCellStyle(15, 2, 15, 6, mStyles.xlStyleKhaki);
                xlWorkSheet.SetCellStyle(16, 2, 16, 6, mStyles.xlStyleLemonChiffon);
                mStyles.xlStyleLemonChiffon.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
                xlWorkSheet.SetCellStyle(11, 1, 11, 1, mStyles.xlStyleLemonChiffon);

                xlWorkSheet.SetCellStyle(18, 1, 18, 6, mStyles.xlStyleKhaki);

                xlWorkSheet.DrawBorderGrid(1, 1, 2, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.DrawBorderGrid(4, 1, 9, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.DrawBorderGrid(11, 1, 16, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.DrawBorderGrid(18, 1, 18, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin);
                xlWorkSheet.AutoFitColumn(1, 6);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E51_Daily_Profit_Totals_per_Category: {ex.Message}", ex);
            }
        }
        public string E53_SeaChefs_InvoicesByDepartureDate()
        {

            RowCounter = 0;

            string client = "";
            string clientName = "";
            string vessel = "";
            string cc = "";

            string prevClient = "";
            string prevVessel = "";
            string prevCC = "";
            decimal[,] totals = new decimal[4, 5];

            try
            {

                xlWorkSheet.SetColumnStyle(1, 18, mStyles.xlStyleText);
                xlWorkSheet.SetColumnStyle(9, 13, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(4, mStyles.xlStyleDate);

                RowCounter = 1;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    client = row[0] == DBNull.Value ? string.Empty : row[0]?.ToString() ?? string.Empty;
                    clientName = row[1] == DBNull.Value ? string.Empty : row[1]?.ToString() ?? string.Empty;
                    vessel = row[2] == DBNull.Value ? string.Empty : row[2]?.ToString() ?? string.Empty;
                    cc = row[3] == DBNull.Value ? string.Empty : row[3]?.ToString() ?? string.Empty;
                    if (RowCounter == 1 || client != prevClient || vessel != prevVessel || cc != prevCC)
                    {
                        if (RowCounter > 1)
                        {
                            // CC Total
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 7, prevCC);
                            xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                            xlWorkSheet.SetCellValue(RowCounter, 9, totals[0, 0]);
                            xlWorkSheet.SetCellValue(RowCounter, 10, totals[0, 1]);
                            xlWorkSheet.SetCellValue(RowCounter, 11, totals[0, 2]);
                            xlWorkSheet.SetCellValue(RowCounter, 12, totals[0, 3]);
                            xlWorkSheet.SetCellValue(RowCounter, 13, totals[0, 4]);
                            xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);
                            for (int t = 0; t < 5; t++) totals[0, t] = 0;

                            if (client != prevClient || vessel != prevVessel)
                            {
                                // Vessel Total
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 7, prevVessel);
                                xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                                xlWorkSheet.SetCellValue(RowCounter, 9, totals[1, 0]);
                                xlWorkSheet.SetCellValue(RowCounter, 10, totals[1, 1]);
                                xlWorkSheet.SetCellValue(RowCounter, 11, totals[1, 2]);
                                xlWorkSheet.SetCellValue(RowCounter, 12, totals[1, 3]);
                                xlWorkSheet.SetCellValue(RowCounter, 13, totals[1, 4]);
                                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);
                                for (int t = 0; t < 5; t++) totals[1, t] = 0;
                            }
                            if (client != prevClient)
                            {
                                // Client Total
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 7, prevClient);
                                xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                                xlWorkSheet.SetCellValue(RowCounter, 9, totals[2, 0]);
                                xlWorkSheet.SetCellValue(RowCounter, 10, totals[2, 1]);
                                xlWorkSheet.SetCellValue(RowCounter, 11, totals[2, 2]);
                                xlWorkSheet.SetCellValue(RowCounter, 12, totals[2, 3]);
                                xlWorkSheet.SetCellValue(RowCounter, 13, totals[2, 4]);
                                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);
                                for (int t = 0; t < 5; t++) totals[2, t] = 0;
                            }
                        }
                        if (RowCounter == 1 || client != prevClient)
                        {
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Period:({mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy})");
                            xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, 18);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, mStyles.xlStyleBisque);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, "Currency");
                            xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[7]));
                            xlWorkSheet.SetCellValue(RowCounter, 3, "Account");
                            xlWorkSheet.SetCellValue(RowCounter, 4, client);
                            xlWorkSheet.SetCellValue(RowCounter, 5, clientName);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 5, mStyles.xlStyleBold);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, "Project");
                            xlWorkSheet.SetCellValue(RowCounter, 2, "Cost Center");
                            xlWorkSheet.SetCellValue(RowCounter, 3, "Invoice");
                            xlWorkSheet.SetCellValue(RowCounter, 4, "Dep Date");
                            xlWorkSheet.SetCellValue(RowCounter, 5, "AL");
                            xlWorkSheet.SetCellValue(RowCounter, 6, "Traveller");
                            xlWorkSheet.SetCellValue(RowCounter, 7, "TKT No");
                            xlWorkSheet.SetCellValue(RowCounter, 8, "Routing");
                            xlWorkSheet.SetCellValue(RowCounter, 9, "Fare");
                            xlWorkSheet.SetCellValue(RowCounter, 10, "Taxes");
                            xlWorkSheet.SetCellValue(RowCounter, 11, "Canc Fee");
                            xlWorkSheet.SetCellValue(RowCounter, 12, "Discount");
                            xlWorkSheet.SetCellValue(RowCounter, 13, "Payable");
                            xlWorkSheet.SetCellValue(RowCounter, 14, "Trip ID");
                            xlWorkSheet.SetCellValue(RowCounter, 15, "Order by");
                            xlWorkSheet.SetCellValue(RowCounter, 16, "Invoice Ref.");
                            xlWorkSheet.SetCellValue(RowCounter, 17, "Reason For travel");
                            xlWorkSheet.SetCellValue(RowCounter, 18, "Passenger ID");
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 18, mStyles.xlStyleBisqueWithBorder);
                        }
                        if (prevClient != client || prevVessel != vessel)
                        {
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, vessel);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 1, mStyles.xlStyleBold);
                        }
                        if (prevClient != client || prevVessel != vessel || prevCC != cc)
                        {
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, cc);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 1, mStyles.xlStyleBold);
                        }

                        prevClient = client;
                        prevVessel = vessel;
                        prevCC = cc;
                    }
                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[2]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDateTime(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToDecimal(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[15]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[16]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[17]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[19]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, "");
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[20]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[21]));
                    for (int iTot = 0; iTot < 4; iTot++)
                    {
                        totals[iTot, 0] += Convert.ToDecimal(row[13]);
                        totals[iTot, 1] += Convert.ToDecimal(row[14]);
                        totals[iTot, 2] += Convert.ToDecimal(row[15]);
                        totals[iTot, 3] += Convert.ToDecimal(row[16]);
                        totals[iTot, 4] += Convert.ToDecimal(row[17]);
                    }
                }
                // CC Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 7, prevCC);
                xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 9, totals[0, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[0, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[0, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[0, 3]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[0, 4]);
                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);

                // Vessel Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 7, prevVessel);
                xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 9, totals[1, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[1, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[1, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[1, 3]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[1, 4]);
                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);

                // Client Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 7, prevClient);
                xlWorkSheet.SetCellValue(RowCounter, 8, "SUBTOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 9, totals[2, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[2, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[2, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[2, 3]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[2, 4]);
                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);

                // Grand Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 8, "TOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 9, totals[3, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[3, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[3, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[3, 3]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[3, 4]);
                xlWorkSheet.SetCellStyle(RowCounter, 7, RowCounter, 13, mStyles.xlStyleBold);

                xlWorkSheet.AutoFitColumn(1, 18);

                xlWorkSheet.SaveAs(FileName);
                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception("E53_SeaChefs_InvoicesByDepartureDate(ReportsNext.ReportsCollection mReport, string FileName)\r\n" + ex.Message, ex);
            }
        }
        public string E54_Client_Statement()
        {

            RowCounter = 2;

            string client = "";
            string clientName = "";
            string vessel = "";

            string prevClient = "";
            string prevClientName = "";
            string prevVessel = "";
            decimal[,] totals = new decimal[3, 4];

            int numCols = 13;
            try
            {
                if (mobjReports.BooleanOption1)
                {
                    xlWorkSheet.SetCellValue(1, 2, "INCLUDING OMIT");
                    xlWorkSheet.SetCellStyle(1, 2, mStyles.xlStyleSandyBrown);
                    numCols = 16;
                }

                xlWorkSheet.SetColumnStyle(1, numCols, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(10, 13, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(10, 16, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(2, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(RowCounter, 1, CompanyNameEnglish);
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, CompanyNameGreek);
                RowCounter++;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    client = row[0] == DBNull.Value ? string.Empty : row[0]?.ToString() ?? string.Empty;
                    clientName = row[1] == DBNull.Value ? string.Empty : row[1]?.ToString() ?? string.Empty;
                    vessel = row[5] == DBNull.Value ? string.Empty : row[5]?.ToString() ?? string.Empty;

                    if (RowCounter == 4 || client != prevClient || vessel != prevVessel)
                    {
                        if (RowCounter > 4)
                        {
                            // Vessel Total
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Total for Vessel {prevVessel}:");
                            xlWorkSheet.SetCellValue(RowCounter, 10, totals[0, 0]);
                            xlWorkSheet.SetCellValue(RowCounter, 11, totals[0, 1]);
                            xlWorkSheet.SetCellValue(RowCounter, 12, totals[0, 2]);
                            xlWorkSheet.SetCellValue(RowCounter, 13, totals[0, 3]);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlBoldWithBorder);
                            totals[0, 0] = totals[0, 1] = totals[0, 2] = totals[0, 3] = 0;

                            if (client != prevClient)
                            {
                                // Client Total
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total Payable {prevClient} {prevClientName}:");
                                xlWorkSheet.SetCellValue(RowCounter, 10, totals[1, 0]);
                                xlWorkSheet.SetCellValue(RowCounter, 11, totals[1, 1]);
                                xlWorkSheet.SetCellValue(RowCounter, 12, totals[1, 2]);
                                xlWorkSheet.SetCellValue(RowCounter, 13, totals[1, 3]);
                                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlBoldWithBorder);
                                totals[1, 0] = totals[1, 1] = totals[1, 2] = totals[1, 3] = 0;
                            }
                        }
                        if (RowCounter == 1 || client != prevClient)
                        {
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Statement Period:({mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy})");
                            xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, numCols);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, mStyles.xlStyleBisque);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Client: {client} {clientName}");
                            xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[2]));
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 5, mStyles.xlBoldWithBorder);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, "P-T Number");
                            xlWorkSheet.SetCellValue(RowCounter, 2, "Inv. Date");
                            xlWorkSheet.SetCellValue(RowCounter, 3, "Vessel");
                            xlWorkSheet.SetCellValue(RowCounter, 4, "Description");
                            xlWorkSheet.SetCellValue(RowCounter, 5, "Pax");
                            xlWorkSheet.SetCellValue(RowCounter, 6, "A/L");
                            xlWorkSheet.SetCellValue(RowCounter, 7, "Routing");
                            xlWorkSheet.SetCellValue(RowCounter, 8, "Flight Date");
                            xlWorkSheet.SetCellValue(RowCounter, 9, "ReasonForTravel");
                            xlWorkSheet.SetCellValue(RowCounter, 10, "Taxes");
                            xlWorkSheet.SetCellValue(RowCounter, 11, "Face Value");
                            xlWorkSheet.SetCellValue(RowCounter, 12, "Discount");
                            xlWorkSheet.SetCellValue(RowCounter, 13, "Net Payable");
                            if (mobjReports.BooleanOption1)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 14, "OMIT");
                                xlWorkSheet.SetCellValue(RowCounter, 15, "ConnectedDocument");
                                xlWorkSheet.SetCellValue(RowCounter, 16, "ConnectedDoc.Amount");
                            }
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlStyleBisqueWithBorder);
                        }
                        if (prevClient != client || prevVessel != vessel)
                        {
                            RowCounter++;
                        }
                        prevClient = client;
                        prevClientName = clientName;
                        prevVessel = vessel;
                    }
                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToDateTime(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToInt32(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToDateTime(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDecimal(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDecimal(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToDecimal(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDecimal(row[15]));
                    if (mobjReports.BooleanOption1)
                    {
                        if (Convert.ToBoolean(row[16]))
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 14, "OMIT");
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlStyleSandyBrown);
                        }
                        xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row[17]));
                        xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[18]));
                    }
                    for (int iTot = 0; iTot <= 2; iTot++)
                    {
                        totals[iTot, 0] += Convert.ToDecimal(row[12]);
                        totals[iTot, 1] += Convert.ToDecimal(row[13]);
                        totals[iTot, 2] += Convert.ToDecimal(row[14]);
                        totals[iTot, 3] += Convert.ToDecimal(row[15]);
                    }
                }

                // Vessel Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total for Vessel {prevVessel}:");
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[0, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[0, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[0, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[0, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlBoldWithBorder);

                // Client Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total Payable {prevClient} {prevClientName}:");
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[1, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[1, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[1, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[1, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlBoldWithBorder);

                // Grand Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, "TOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 10, totals[2, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 11, totals[2, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 12, totals[2, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 13, totals[2, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, numCols, mStyles.xlBoldWithBorder);

                xlWorkSheet.AutoFitColumn(1, numCols);
                xlWorkSheet.SaveAs(FileName);

                return FileName;
            }
            catch (Exception ex)
            {
                throw new Exception($"E54_Client_Statement(ReportsNext.ReportsCollection mobjReports, string FileName)\r\n{ex.Message}", ex);
            }
        }
        public string E56_ClientsPerGroup()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 4, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(5, 6, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(1, 1, "Client Code");
                xlWorkSheet.SetCellValue(1, 2, "Client Name");
                xlWorkSheet.SetCellValue(1, 3, "Client Group");
                xlWorkSheet.SetCellValue(1, 4, "Agent Group");
                xlWorkSheet.SetCellValue(1, 5, "Last Transaction");
                xlWorkSheet.SetCellValue(1, 6, "Date Created");

                xlWorkSheet.SetCellStyle(1, 1, 1, 6, mStyles.xlStyleHeader);

                RowCounter = 1;
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(mdsDataSet.Tables[0].Rows[i][0]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(mdsDataSet.Tables[0].Rows[i][1]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(mdsDataSet.Tables[0].Rows[i][2]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(mdsDataSet.Tables[0].Rows[i][3]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDateTime(mdsDataSet.Tables[0].Rows[i][5]));


                    if (mdsDataSet.Tables[0].Rows[i][4] == DBNull.Value ||
                        ((mdsDataSet.Tables[0].Rows[i][4] is DateTime lastTrans) &&
                         ((DateTime.Now.Year - lastTrans.Year) * 12 + DateTime.Now.Month - lastTrans.Month > 24)))
                    {
                        xlWorkSheet.SetCellStyle(RowCounter, 5, mStyles.xlStyleSandyBrown);
                    }
                    else
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDateTime(mdsDataSet.Tables[0].Rows[i][4]));
                    }
                }

                xlWorkSheet.AutoFitColumn(1, 6);
                xlWorkSheet.SaveAs(FileName);

                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception($"E56_ClientsPerGroup\r\n{ex.Message}");
            }
        }
        public string E60_Report_For_Lowest_Class()
        {
             RowCounter = 1;
            string client, clientName, vessel;
            int columnCount = 23;

            try
            {


                xlWorkSheet.SetColumnStyle(1, columnCount, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(15, 18, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(5, 6, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(14, mStyles.xlStyleDate);

                xlWorkSheet.FreezePanes(1, 0);

                // Header
                xlWorkSheet.SetCellValue(RowCounter, 1, "Client");
                xlWorkSheet.SetCellValue(RowCounter, 2, "Client Name");
                xlWorkSheet.SetCellValue(RowCounter, 3, "Vessel");
                xlWorkSheet.SetCellValue(RowCounter, 4, "P-T Number");
                xlWorkSheet.SetCellValue(RowCounter, 5, "Inv. Date");
                xlWorkSheet.SetCellValue(RowCounter, 6, "Booking Date");
                xlWorkSheet.SetCellValue(RowCounter, 7, "Description");
                xlWorkSheet.SetCellValue(RowCounter, 8, "Class Of Service");
                xlWorkSheet.SetCellValue(RowCounter, 9, "Cabin");
                xlWorkSheet.SetCellValue(RowCounter, 10, "Tkt");
                xlWorkSheet.SetCellValue(RowCounter, 11, "Pax");
                xlWorkSheet.SetCellValue(RowCounter, 12, "A/L");
                xlWorkSheet.SetCellValue(RowCounter, 13, "Routing");
                xlWorkSheet.SetCellValue(RowCounter, 14, "Flight Date");
                xlWorkSheet.SetCellValue(RowCounter, 15, "Fare");
                xlWorkSheet.SetCellValue(RowCounter, 16, "Taxes");
                xlWorkSheet.SetCellValue(RowCounter, 17, "Discount");
                xlWorkSheet.SetCellValue(RowCounter, 18, "Total Payable");
                xlWorkSheet.SetCellValue(RowCounter, 19, "Fare Basis");
                xlWorkSheet.SetCellValue(RowCounter, 20, "Transaction Type");
                xlWorkSheet.SetCellValue(RowCounter, 21, "VOID");
                xlWorkSheet.SetCellValue(RowCounter, 22, "Connected Document");
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlStyleBisqueWithBorder);

                // Data rows
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    client = row[0] == DBNull.Value ? string.Empty : row[0]?.ToString() ?? string.Empty;
                    clientName = row[1] == DBNull.Value ? string.Empty : row[1]?.ToString() ?? string.Empty;
                    vessel = row[5] == DBNull.Value ? string.Empty : row[5]?.ToString() ?? string.Empty;

                    RowCounter++;
                    xlWorkSheet.SetCellValue(RowCounter, 1, client);// Client
                    xlWorkSheet.SetCellValue(RowCounter, 2, clientName);// Client Name
                    xlWorkSheet.SetCellValue(RowCounter, 3, vessel);// Vessel
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3]));// P-T Number
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToDateTime(row[4]));// Inv. Date
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDateTime(row["EntryDate"]));// Booking Date
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6]));// Description
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[16]));// Class Of Service
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[17]));// Cabin
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[18]));// Tkt
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToInt32(row[7]));// Pax
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[8]));// A/L
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[9]));// Routing
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDateTime(row[10]));// Flight Date
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDecimal(row[12]));// Fare
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[11]));// Taxes
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[13]));// Discount
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToDecimal(row[14]));// Total Payable
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row["FareBasis"]));// Fare Basis
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row["TransactionType"]));// Transaction Type
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row["Void"]));// VOID
                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row["ConnectedDocument"]));// Connected Document

                    if (Convert.ToInt32(row["Omit"]) == 1)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 23, "OMIT");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlStyleSandyBrown);
                    }
                }

                xlWorkSheet.AutoFitColumn(1, columnCount);
                xlWorkSheet.SaveAs(FileName);

                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception($"E60_Report_For_Lowest_Class\r\n{ex.Message}");
            }
        }
        public string E63_AirTicketSalesTemenos()
        {

            RowCounter = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 20, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(6, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(4, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(10, 11, mStyles.xlStyleDate);

                xlWorkSheet.SetCellValue(1, 1, "Passenger");
                xlWorkSheet.SetCellValue(1, 2, "Inv Code");
                xlWorkSheet.SetCellValue(1, 3, "Inv Number");
                xlWorkSheet.SetCellValue(1, 4, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 5, "Vessel");
                xlWorkSheet.SetCellValue(1, 6, "Net Payable");
                xlWorkSheet.SetCellValue(1, 7, "Transaction Type");
                xlWorkSheet.SetCellValue(1, 8, "Airline");
                xlWorkSheet.SetCellValue(1, 9, "Routing");
                xlWorkSheet.SetCellValue(1, 10, "Departure Date");
                xlWorkSheet.SetCellValue(1, 11, "Arrival Date");
                xlWorkSheet.SetCellValue(1, 12, "PNR");
                xlWorkSheet.SetCellValue(1, 13, "BookedBy");
                xlWorkSheet.SetCellValue(1, 14, "ReasonForTravel");
                xlWorkSheet.SetCellValue(1, 15, "CostCentre");
                xlWorkSheet.SetCellValue(1, 16, "Employee ID");
                xlWorkSheet.SetCellValue(1, 17, "Project LIST A");
                xlWorkSheet.SetCellValue(1, 18, "Employee Type");
                xlWorkSheet.SetCellValue(1, 19, "Entity Code");
                xlWorkSheet.SetCellValue(1, 20, "Project LIST B");

                xlWorkSheet.SetCellStyle(1, 1, 1, 20, mStyles.xlStyleHeader);

                RowCounter = 1;
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;

                    var row = mdsDataSet.Tables[0].Rows[i];

                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row["Passenger"]));
                    if (Convert.ToInt32(row["InvNumber"]) != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row["InvCode"]));
                        xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row["InvNumber"]));
                        xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToDateTime(row["InvoiceDate"]));
                    }
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row["Vessel"]));
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToDecimal(row["NetPayable"]));
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row["TransactionType"]));
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row["TicketingAirline"]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row["Routing"]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToDateTime(row["DepartureDate"]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToDateTime(row["ArrivalDate"]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row["PNR"]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row["01-BookedBy"]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row["04-ReasonForTravel"]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToString(row["05-CostCentre"]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row["12-Passenger ID"]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row["18-REF1"]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row["20-REF3"]));
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row["23-REF6"]));
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row["19-REF2"]));

                    if (!string.IsNullOrEmpty(row["Omit"]?.ToString()))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 20, mStyles.xlStyleSandyBrown);
                    if (!string.IsNullOrEmpty(row["Void"]?.ToString()))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 20, mStyles.xlStyleGrayItalic);
                    if (row["ActionType"]?.ToString() == "Refund")
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 20, mStyles.xlStyleRedFont);
                }

                xlWorkSheet.AutoFitColumn(1, 20);
                xlWorkSheet.SaveAs(FileName);


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return ($"File saved: {FileName}");
        }
        public string E64_LowestClasses()
        {
             RowCounter = 2;
            string? client, clientName, vessel, bookingClass, classRank, itin, airTicketType, isExchangedText, prevClient = "", prevClientName = "", prevVessel = "";
            int classLevel;
            long airTicketTypeId;
            bool isExchanged, hasRelatedTicket;
            decimal[,] totals = new decimal[3, 4];

            try
            {
                int columnCount = mobjReports.BooleanOption1 ? 20 : 19;

                xlWorkSheet.SetColumnStyle(1, columnCount, mStyles.xlStyleText);
                xlWorkSheet.SetColumnStyle(14, 17, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(2, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(13, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(5, 6, mStyles.xlCentred);

                // Header rows
                xlWorkSheet.SetCellValue(RowCounter, 1, CompanyNameEnglish);
                xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, columnCount);
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, CompanyNameGreek);
                xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, columnCount);
                RowCounter++;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    client = row[0] == DBNull.Value ? string.Empty : row[0]?.ToString() ?? string.Empty;
                    clientName = row[1] == DBNull.Value ? string.Empty : row[1]?.ToString() ?? string.Empty;
                    vessel = row[5] == DBNull.Value ? string.Empty : row[5]?.ToString() ?? string.Empty;
                    bookingClass = row[16] == DBNull.Value ? string.Empty : row[16]?.ToString() ?? string.Empty;
                    classRank = row[26] == DBNull.Value ? string.Empty : row[26]?.ToString() ?? string.Empty;
                    classLevel = GetClassLevel(bookingClass, classRank);
                    itin = row[9] == DBNull.Value ? string.Empty : row[9]?.ToString() ?? string.Empty;
                    airTicketTypeId = Convert.ToInt64(row[19]);
                    airTicketType = row[29] == DBNull.Value ? string.Empty : row[29]?.ToString() ?? string.Empty;
                    isExchanged = Convert.ToBoolean(row[31]);
                    isExchangedText = isExchanged ? $"Reissue {row[32]}" : "";
                    hasRelatedTicket = Convert.ToBoolean(row[33]);
                    if (hasRelatedTicket)
                    {
                        if (!string.IsNullOrEmpty(isExchangedText))
                            isExchangedText += " - ";
                        isExchangedText += $"Related {row[34]}";
                    }

                    if (RowCounter == 4 || client != prevClient || vessel != prevVessel)
                    {
                        if (RowCounter > 4)
                        {
                            // Vessel Total
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Total for Vessel {prevVessel}:");
                            xlWorkSheet.SetCellValue(RowCounter, 14, totals[0, 0]);
                            xlWorkSheet.SetCellValue(RowCounter, 15, totals[0, 1]);
                            xlWorkSheet.SetCellValue(RowCounter, 16, totals[0, 2]);
                            xlWorkSheet.SetCellValue(RowCounter, 17, totals[0, 3]);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlBoldWithBorder);
                            totals[0, 0] = totals[0, 1] = totals[0, 2] = totals[0, 3] = 0;

                            if (client != prevClient)
                            {
                                // Client Total
                                RowCounter++;
                                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total Payable {prevClient} {prevClientName}:");
                                xlWorkSheet.SetCellValue(RowCounter, 14, totals[1, 0]);
                                xlWorkSheet.SetCellValue(RowCounter, 15, totals[1, 1]);
                                xlWorkSheet.SetCellValue(RowCounter, 16, totals[1, 2]);
                                xlWorkSheet.SetCellValue(RowCounter, 17, totals[1, 3]);
                                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlBoldWithBorder);
                                totals[1, 0] = totals[1, 1] = totals[1, 2] = totals[1, 3] = 0;
                            }
                        }
                        if (RowCounter == 1 || client != prevClient)
                        {
                            RowCounter += 2;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Statement Period:({mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy})");
                            xlWorkSheet.MergeWorksheetCells(RowCounter, 1, RowCounter, columnCount);
                            xlWorkSheet.SetCellStyle(RowCounter, 1, mStyles.xlStyleBisque);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, $"Client: {client} {clientName}");
                            xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToString(row[2]));
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlCyanWithBorder);
                            RowCounter++;
                            xlWorkSheet.SetCellValue(RowCounter, 1, "P-T Number");
                            xlWorkSheet.SetCellValue(RowCounter, 2, "Inv. Date");
                            xlWorkSheet.SetCellValue(RowCounter, 3, "Vessel");
                            xlWorkSheet.SetCellValue(RowCounter, 4, "Description");
                            xlWorkSheet.SetCellValue(RowCounter, 5, "Booking Class");
                            xlWorkSheet.SetCellValue(RowCounter, 6, "Class Rank");
                            xlWorkSheet.SetCellValue(RowCounter, 7, "Marine");
                            xlWorkSheet.SetCellValue(RowCounter, 8, "Class Of Service");
                            xlWorkSheet.SetCellValue(RowCounter, 9, "Tkt");
                            xlWorkSheet.SetCellValue(RowCounter, 10, "Pax");
                            xlWorkSheet.SetCellValue(RowCounter, 11, "A/L");
                            xlWorkSheet.SetCellValue(RowCounter, 12, "Routing");
                            xlWorkSheet.SetCellValue(RowCounter, 13, "Flight Date");
                            xlWorkSheet.SetCellValue(RowCounter, 14, "Fare");
                            xlWorkSheet.SetCellValue(RowCounter, 15, "Taxes");
                            xlWorkSheet.SetCellValue(RowCounter, 16, "Discount");
                            xlWorkSheet.SetCellValue(RowCounter, 17, "Total Payable");
                            xlWorkSheet.SetCellValue(RowCounter, 18, "Lowest Fare");
                            if (mobjReports.BooleanOption1)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 19, "Fare Basis");
                                xlWorkSheet.SetCellValue(RowCounter, 20, "Notes");
                            }
                            else
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 19, "Notes");
                            }
                            xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlStyleBisqueWithBorder);
                        }
                        if (prevClient != client || prevVessel != vessel)
                        {
                            RowCounter++;
                        }
                        prevClient = client;
                        prevClientName = clientName;
                        prevVessel = vessel;
                    }
                    RowCounter++;

                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[3]));
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToDateTime(row[4]));
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[5]));
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[6]));
                    xlWorkSheet.SetCellValue(RowCounter, 5, bookingClass);

                    if (airTicketTypeId != 323)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 6, airTicketType);
                        xlWorkSheet.SetCellStyle(RowCounter, 6, mStyles.xlStyleYellowBold);
                    }
                    else
                    {
                        if (itin != null && itin.Length > 6)
                        {
                            if (itin[..3] == itin.Substring(itin.Length - 3, 3))
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 6, "RT");
                                xlWorkSheet.SetCellStyle(RowCounter, 6, mStyles.xlStyleYellowBold);
                            }
                            else if (classLevel > 0)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, 6, classLevel);
                                if (classLevel <= 2)
                                    xlWorkSheet.SetCellStyle(RowCounter, 6, mStyles.xlStyleLightGreen);
                            }
                        }
                        if (Convert.ToInt32(row[27]) == 1)
                        {
                            xlWorkSheet.SetCellValue(RowCounter, 7, "Marine");
                        }
                    }

                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[17]));
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToString(row[18]));
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToInt32(row[7]));
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[8]));
                    xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[9]));
                    xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToDateTime(row[10]));
                    xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToDecimal(row[12]));
                    xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDecimal(row[11]));
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToDecimal(row[13]));
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[14]));
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[35]));
                    if (mobjReports.BooleanOption1)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row["FareBasis"]));
                        xlWorkSheet.SetCellValue(RowCounter, 20, isExchangedText);
                    }
                    else
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 19, isExchangedText);
                    }
                    for (int iTot = 0; iTot <= 2; iTot++)
                    {
                        totals[iTot, 0] += Convert.ToDecimal(row[12]);
                        totals[iTot, 1] += Convert.ToDecimal(row[11]);
                        totals[iTot, 2] += Convert.ToDecimal(row[13]);
                        totals[iTot, 3] += Convert.ToDecimal(row[14]);
                    }
                }

                // Vessel Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total for Vessel {prevVessel}:");
                xlWorkSheet.SetCellValue(RowCounter, 14, totals[0, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 15, totals[0, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 16, totals[0, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 17, totals[0, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlBoldWithBorder);

                // Client Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, $"Total Payable {prevClient} {prevClientName}:");
                xlWorkSheet.SetCellValue(RowCounter, 14, totals[1, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 15, totals[1, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 16, totals[1, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 17, totals[1, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlBoldWithBorder);

                // Grand Total
                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, "TOTALS:");
                xlWorkSheet.SetCellValue(RowCounter, 14, totals[2, 0]);
                xlWorkSheet.SetCellValue(RowCounter, 15, totals[2, 1]);
                xlWorkSheet.SetCellValue(RowCounter, 16, totals[2, 2]);
                xlWorkSheet.SetCellValue(RowCounter, 17, totals[2, 3]);
                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, columnCount, mStyles.xlBoldWithBorder);

                xlWorkSheet.AutoFitColumn(1, columnCount);
                xlWorkSheet.SaveAs(FileName);

                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private int GetClassLevel(string? bookingClass, string? classRank)
        {
            if (bookingClass == null || classRank == null) return 0;
            var cc = classRank.Split(['|'], StringSplitOptions.RemoveEmptyEntries);
            int ret = -1;
            if (cc.Length > 1 && cc[1].Contains(bookingClass))
            {
                ret = 1;
            }
            else if (cc.Length > 0)
            {
                ret = cc[0].IndexOf(bookingClass) + 1;
            }
            return ret;
        }
        public string E65_OpsSales()
        {
            RowCounter = 1;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales");
                xlWorkSheet.FreezePanes(1, 0);

                xlWorkSheet.SetColumnStyle(1, 39, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(17, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(36, 38, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(15, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(28, 29, mStyles.xlStyleDate);

                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(39, mStyles.xlStyleInteger);

                // Header
                xlWorkSheet.SetCellValue(1, 1, "Issue Date");
                xlWorkSheet.SetCellValue(1, 2, "Client Code");
                xlWorkSheet.SetCellValue(1, 3, "Client Name");
                xlWorkSheet.SetCellValue(1, 4, "Omit");
                xlWorkSheet.SetCellValue(1, 5, "Void");
                xlWorkSheet.SetCellValue(1, 6, "PNR");
                xlWorkSheet.SetCellValue(1, 7, "Ticket Number");
                xlWorkSheet.SetCellValue(1, 8, "Passenger");
                xlWorkSheet.SetCellValue(1, 9, "Pax Count");
                xlWorkSheet.SetCellValue(1, 10, "Product Type");
                xlWorkSheet.SetCellValue(1, 11, "Action Type");
                xlWorkSheet.SetCellValue(1, 12, "Inv Code");
                xlWorkSheet.SetCellValue(1, 13, "Inv Series");
                xlWorkSheet.SetCellValue(1, 14, "Inv Number");
                xlWorkSheet.SetCellValue(1, 15, "Invoice Date");
                xlWorkSheet.SetCellValue(1, 16, "Vessel");
                xlWorkSheet.SetCellValue(1, 17, "Net Payable");
                xlWorkSheet.SetCellValue(1, 18, "Verified");
                xlWorkSheet.SetCellValue(1, 19, "Remarks");
                xlWorkSheet.SetCellValue(1, 20, "Transaction Type");
                xlWorkSheet.SetCellValue(1, 21, "RegNr");
                xlWorkSheet.SetCellValue(1, 22, "Ticketing Airline");
                xlWorkSheet.SetCellValue(1, 23, "Routing");
                xlWorkSheet.SetCellValue(1, 24, "SalesPerson");
                xlWorkSheet.SetCellValue(1, 25, "Issuing Agent");
                xlWorkSheet.SetCellValue(1, 26, "Creator Agent");
                xlWorkSheet.SetCellValue(1, 27, "Reference");
                xlWorkSheet.SetCellValue(1, 28, "Departure Date");
                xlWorkSheet.SetCellValue(1, 29, "Arrival Date");
                xlWorkSheet.SetCellValue(1, 30, "Connected Document");
                xlWorkSheet.SetCellValue(1, 31, "Pax Remarks");
                xlWorkSheet.SetCellValue(1, 32, "Doc Status ID");
                xlWorkSheet.SetCellValue(1, 33, "Cancels Docs");
                xlWorkSheet.SetCellValue(1, 34, "Sertvices Description");
                xlWorkSheet.SetCellValue(1, 35, "Client Team");
                xlWorkSheet.SetCellValue(1, 36, "Sell");
                xlWorkSheet.SetCellValue(1, 37, "Buy");
                xlWorkSheet.SetCellValue(1, 38, "Profit");
                xlWorkSheet.SetCellValue(1, 39, "PaxCount+-");
                xlWorkSheet.SetCellStyle(1, 1, 1, 39, mStyles.xlStyleHeader);

                // Data rows
                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToDateTime(row[0])); // Issue Date
                    xlWorkSheet.SetCellValue(RowCounter, 2, Convert.ToString(row[1])); // Client Code
                    xlWorkSheet.SetCellValue(RowCounter, 3, Convert.ToString(row[2])); // Client Name
                    xlWorkSheet.SetCellValue(RowCounter, 4, Convert.ToString(row[3])); // Omit
                    xlWorkSheet.SetCellValue(RowCounter, 5, Convert.ToString(row[4])); // Void
                    xlWorkSheet.SetCellValue(RowCounter, 6, Convert.ToString(row[5])); // PNR
                    xlWorkSheet.SetCellValue(RowCounter, 7, Convert.ToString(row[6])); // Ticket Number
                    xlWorkSheet.SetCellValue(RowCounter, 8, Convert.ToString(row[7])); // Passenger
                    xlWorkSheet.SetCellValue(RowCounter, 9, Convert.ToInt32(row[8])); // Pax Count
                    xlWorkSheet.SetCellValue(RowCounter, 10, Convert.ToString(row[9])); // Product Type
                    xlWorkSheet.SetCellValue(RowCounter, 11, Convert.ToString(row[10])); // Action Type
                    if (Convert.ToInt32(row[13]) != 0)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 12, Convert.ToString(row[11])); // Inv Code
                        xlWorkSheet.SetCellValue(RowCounter, 13, Convert.ToString(row[12])); // Inv Series
                        xlWorkSheet.SetCellValue(RowCounter, 14, Convert.ToString(row[13])); // Inv Number
                        xlWorkSheet.SetCellValue(RowCounter, 15, Convert.ToDateTime(row[14])); // Invoice Date
                    }
                    xlWorkSheet.SetCellValue(RowCounter, 16, Convert.ToString(row[15])); // Vessel
                    xlWorkSheet.SetCellValue(RowCounter, 17, Convert.ToDecimal(row[16])); // Net Payable
                    xlWorkSheet.SetCellValue(RowCounter, 18, Convert.ToString(row[17])); // Verified
                    xlWorkSheet.SetCellValue(RowCounter, 19, Convert.ToString(row[18])); // Remarks
                    xlWorkSheet.SetCellValue(RowCounter, 20, Convert.ToString(row[19])); // Transaction Type
                    xlWorkSheet.SetCellValue(RowCounter, 21, Convert.ToString(row[20])); // RegNr
                    xlWorkSheet.SetCellValue(RowCounter, 22, Convert.ToString(row[21])); // Ticketing Airline
                    xlWorkSheet.SetCellValue(RowCounter, 23, Convert.ToString(row[22])); // Routing
                    xlWorkSheet.SetCellValue(RowCounter, 24, Convert.ToString(row[23])); // SalesPerson
                    xlWorkSheet.SetCellValue(RowCounter, 25, Convert.ToString(row[24])); // Issuing Agent
                    xlWorkSheet.SetCellValue(RowCounter, 26, Convert.ToString(row[25])); // Creator Agent
                    xlWorkSheet.SetCellValue(RowCounter, 27, Convert.ToString(row[26])); // Reference
                    xlWorkSheet.SetCellValue(RowCounter, 28, Convert.ToDateTime(row[27])); // Departure Date
                    xlWorkSheet.SetCellValue(RowCounter, 29, Convert.ToDateTime(row[28])); // Arrival Date
                    xlWorkSheet.SetCellValue(RowCounter, 30, Convert.ToString(row[29])); // Connected Document
                    xlWorkSheet.SetCellValue(RowCounter, 31, Convert.ToString(row[30])); // Pax Remarks
                    xlWorkSheet.SetCellValue(RowCounter, 32, Convert.ToInt32(row[31])); // Doc Status ID
                    xlWorkSheet.SetCellValue(RowCounter, 33, Convert.ToString(row[32])); // Cancels Docs
                    xlWorkSheet.SetCellValue(RowCounter, 34, Convert.ToString(row[33])); // Sertvices Description
                    xlWorkSheet.SetCellValue(RowCounter, 35, Convert.ToString(row[34])); // Client Team
                    xlWorkSheet.SetCellValue(RowCounter, 36, Convert.ToDecimal(row[35])); // Sell
                    xlWorkSheet.SetCellValue(RowCounter, 37, Convert.ToDecimal(row[36])); // Buy
                    xlWorkSheet.SetCellValue(RowCounter, 38, Convert.ToDecimal(row[37])); // Profit
                    xlWorkSheet.SetCellValue(RowCounter, 39, Convert.ToInt32(row[38])); // PaxCount+-


                    // Cancelled/Cancelled Docs
                    if (Convert.ToInt32(row[31]) == 43)
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 33, "Cancelled");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 39, mStyles.xlStyleItalic);
                    }
                    else if (!string.IsNullOrEmpty(row[32]?.ToString()))
                    {
                        xlWorkSheet.SetCellValue(RowCounter, 33, $"Cancels {row[32]}");
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 39, mStyles.xlStyleItalic);
                    }


                    // Omit/Void/Refund styling
                    if (!string.IsNullOrEmpty(row[3]?.ToString()))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 39, mStyles.xlStyleSandyBrown);
                    if (!string.IsNullOrEmpty(row[4]?.ToString()))
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 39, mStyles.xlStyleGrayItalic);
                    if (row[10]?.ToString() == "Refund")
                        xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 39, mStyles.xlStyleRedFont);
                }

                xlWorkSheet.AutoFitColumn(1, 39);
                xlWorkSheet.SaveAs(FileName);
                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string E66_PurchasesPerAirline()
        {
            RowCounter = 0;
            int headerRow = 0;
            decimal[] totals = new decimal[19];

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Purchases");

                xlWorkSheet.SetColumnStyle(1, 1, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(2, 4, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(6, 8, mStyles.xlStyleDecimal);
                xlWorkSheet.SetColumnStyle(10, 13, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(5, mStyles.xlStyleInteger);
                xlWorkSheet.SetColumnStyle(9, mStyles.xlStyleInteger);

                RowCounter = 2;
                xlWorkSheet.SetCellValue(RowCounter, 2, $"SALES DATA REPORT - ATHENS {mobjReports.Date1From:dd/MM/yyyy} - {mobjReports.Date1To:dd/MM/yyyy}");
                xlWorkSheet.SetCellStyle(RowCounter, 2, RowCounter, 2, mStyles.xlStyleTitle);
                xlWorkSheet.MergeWorksheetCells(RowCounter, 2, RowCounter, 13);
                RowCounter += 2;

                xlWorkSheet.SetCellValue(RowCounter, 1, "Airline");
                xlWorkSheet.SetCellValue(RowCounter, 2, "Net CY");
                xlWorkSheet.SetCellValue(RowCounter, 3, "Fuel CY");
                xlWorkSheet.SetCellValue(RowCounter, 4, "Net+Fuel CY");
                xlWorkSheet.SetCellValue(RowCounter, 5, "Cpns CY");
                xlWorkSheet.SetCellValue(RowCounter, 6, "Net PY");
                xlWorkSheet.SetCellValue(RowCounter, 7, "Fuel PY");
                xlWorkSheet.SetCellValue(RowCounter, 8, "Net+Fuel PY");
                xlWorkSheet.SetCellValue(RowCounter, 9, "Cpns PY");
                xlWorkSheet.SetCellValue(RowCounter, 10, "Index Net");
                xlWorkSheet.SetCellValue(RowCounter, 11, "Index Fuel");
                xlWorkSheet.SetCellValue(RowCounter, 12, "Index Net+Fuel");
                xlWorkSheet.SetCellValue(RowCounter, 13, "Index Cpns");

                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 13, mStyles.xlStyleHeader);
                xlWorkSheet.FreezePanes(RowCounter, 0);
                headerRow = RowCounter;

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    RowCounter++;
                    var row = mdsDataSet.Tables[0].Rows[i];
                    xlWorkSheet.SetCellValue(RowCounter, 1, Convert.ToString(row[0]));
                    for (int ii = 0; ii <= 2; ii++)
                    {
                        for (int j = 0; j <= 3; j++)
                        {
                            int colIndex = ii * 6 + j + 1;
                            int cellIndex = ii * 4 + j + 2;
                            if (row[colIndex] != DBNull.Value)
                            {
                                xlWorkSheet.SetCellValue(RowCounter, cellIndex, Convert.ToDecimal(row[colIndex]));
                                totals[colIndex] += Convert.ToDecimal(row[colIndex]);
                            }
                        }
                    }
                }

                // Sort by Cpns CY (6th column) and Net CY (2nd column), descending
                xlWorkSheet.Sort(headerRow + 1, 1, RowCounter, 13, 6, false);
                xlWorkSheet.Sort(headerRow + 1, 1, RowCounter, 13, 2, false);

                RowCounter++;
                xlWorkSheet.SetCellValue(RowCounter, 1, "Totals");
                for (int ii = 0; ii <= 1; ii++)
                {
                    for (int j = 0; j <= 3; j++)
                    {
                        int colIndex = ii * 6 + j + 1;
                        int cellIndex = ii * 4 + j + 2;
                        xlWorkSheet.SetCellValue(RowCounter, cellIndex, totals[colIndex]);
                    }
                }

                if (totals[7] != 0) xlWorkSheet.SetCellValue(RowCounter, 10, totals[1] / totals[7] * 100);
                if (totals[8] != 0) xlWorkSheet.SetCellValue(RowCounter, 11, totals[2] / totals[8] * 100);
                if (totals[9] != 0) xlWorkSheet.SetCellValue(RowCounter, 12, totals[3] / totals[9] * 100);
                if (totals[10] != 0) xlWorkSheet.SetCellValue(RowCounter, 13, totals[4] / totals[10] * 100);

                xlWorkSheet.SetCellStyle(RowCounter, 1, RowCounter, 13, mStyles.xlStyleHeader);
                xlWorkSheet.SetCellStyle(headerRow, 1, RowCounter, 1, mStyles.xlStyleHeader);

                xlWorkSheet.DrawBorder(headerRow, 2, RowCounter, 5, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, System.Drawing.Color.Black);
                xlWorkSheet.DrawBorder(headerRow, 6, RowCounter, 9, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, System.Drawing.Color.Black);
                xlWorkSheet.DrawBorder(headerRow, 10, RowCounter, 13, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, System.Drawing.Color.Black);

                xlWorkSheet.AutoFitColumn(1, 13);

                xlWorkSheet.SaveAs(FileName);
                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string E67_Columbia()
        {
            var rowCounter = 0;
            const int ColumnCount = 13;
            string oldClientCode = "";
            decimal totalPayable = 0m;
            string? routing = "";

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Statement");
                xlWorkSheet.FreezePanes(1, 0);


                xlWorkSheet.SetColumnStyle(1, ColumnCount, mStyles.xlStyleText);

                xlWorkSheet.SetColumnStyle(12, mStyles.xlStyleDecimal);

                xlWorkSheet.SetColumnStyle(1, 3, mStyles.xlStyleDate);
                xlWorkSheet.SetColumnStyle(8, mStyles.xlStyleDate);

                rowCounter++;

                xlWorkSheet.SetCellValue(rowCounter, 1, "Booking Date");
                xlWorkSheet.SetCellValue(rowCounter, 2, "Ticket Date");
                xlWorkSheet.SetCellValue(rowCounter, 3, "Invoice Date");
                xlWorkSheet.SetCellValue(rowCounter, 4, "Booked By");
                xlWorkSheet.SetCellValue(rowCounter, 5, "Department");
                xlWorkSheet.SetCellValue(rowCounter, 6, "Vessel");
                xlWorkSheet.SetCellValue(rowCounter, 7, "Invoice Number");
                xlWorkSheet.SetCellValue(rowCounter, 8, "Departure Date");
                xlWorkSheet.SetCellValue(rowCounter, 9, "Destination");
                xlWorkSheet.SetCellValue(rowCounter, 10, "Passenger");
                xlWorkSheet.SetCellValue(rowCounter, 11, "Rank");
                xlWorkSheet.SetCellValue(rowCounter, 12, "Net Payable");
                xlWorkSheet.SetCellValue(rowCounter, 13, "Reason For Travel");
                xlWorkSheet.SetCellStyle(rowCounter, 1, rowCounter, ColumnCount, mStyles.xlStyleHeader);

                for (int i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    var row = mdsDataSet.Tables[0].Rows[i];
                    string? clientCode = row["ClientCode"] == DBNull.Value ? string.Empty : row["ClientCode"]?.ToString() ?? string.Empty;

                    if (string.IsNullOrEmpty(oldClientCode) || oldClientCode != clientCode)
                    {
                        if (!string.IsNullOrEmpty(oldClientCode))
                        {
                            rowCounter++;
                            xlWorkSheet.SetCellValue(rowCounter, 1, "Total Payable:");
                            xlWorkSheet.SetCellValue(rowCounter, 12, totalPayable);
                            xlWorkSheet.SetCellStyle(rowCounter, 1, rowCounter, ColumnCount, mStyles.xlStyleHeader);
                            totalPayable = 0m;
                            rowCounter += 2;
                        }
                        rowCounter++;
                        xlWorkSheet.SetCellValue(rowCounter, 1, $"{clientCode}:{row["ClientName"]}");
                        xlWorkSheet.SetCellStyle(rowCounter, 1, rowCounter, ColumnCount, mStyles.xlStyleTitleBold);
                        xlWorkSheet.MergeWorksheetCells(rowCounter, 1, rowCounter, ColumnCount);
                        oldClientCode = clientCode;
                    }

                    rowCounter++;
                    xlWorkSheet.SetCellValue(rowCounter, 1, Convert.ToDateTime(row["PNRCreationDate"]));
                    xlWorkSheet.SetCellValue(rowCounter, 2, Convert.ToDateTime(row["TicketDate"]));
                    xlWorkSheet.SetCellValue(rowCounter, 3, Convert.ToDateTime(row["InvoiceDate"]));
                    xlWorkSheet.SetCellValue(rowCounter, 4, Convert.ToString(row["BookedBy"]));
                    xlWorkSheet.SetCellValue(rowCounter, 5, Convert.ToString(row["Office"]));
                    xlWorkSheet.SetCellValue(rowCounter, 6, Convert.ToString(row["Vessel"]));
                    xlWorkSheet.SetCellValue(rowCounter, 7, Convert.ToString(row["InvoiceNumber"]));
                    xlWorkSheet.SetCellValue(rowCounter, 8, Convert.ToDateTime(row["DepartureDate"]));

                    if (!string.IsNullOrEmpty(row["AirportName"].ToString()) || !string.IsNullOrEmpty(row["CityName"].ToString()))
                    {
                        if (!string.IsNullOrEmpty(row["CityName"].ToString()))
                            routing = Convert.ToString(row["CityName"].ToString());
                        else
                            routing = $"{row["AirportName"]}/{row["CityName"]}";
                    }
                    else
                    {
                        routing = Convert.ToString(row["Routing"]);
                    }
                    xlWorkSheet.SetCellValue(rowCounter, 9, routing);
                    xlWorkSheet.SetCellValue(rowCounter, 10, Convert.ToString(row["Passenger"]));
                    xlWorkSheet.SetCellValue(rowCounter, 11, Convert.ToString(row["Rank"]));
                    xlWorkSheet.SetCellValue(rowCounter, 12, Convert.ToDecimal(row["NetPayable"]));
                    xlWorkSheet.SetCellValue(rowCounter, 13, Convert.ToString(row["ReasonForTravel"]));

                    if (decimal.TryParse(row["NetPayable"].ToString(), out decimal netPayable))
                        totalPayable += netPayable;
                }

                rowCounter++;
                xlWorkSheet.SetCellValue(rowCounter, 1, "Total Payable:");
                xlWorkSheet.SetCellValue(rowCounter, 12, totalPayable);
                xlWorkSheet.SetCellStyle(rowCounter, 1, rowCounter, ColumnCount, mStyles.xlStyleHeader);
                xlWorkSheet.AutoFitColumn(1, ColumnCount);

                xlWorkSheet.SaveAs(FileName);
                return $"File saved: {FileName}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string E68_GDSImportedPendingItems()
        {
            try
            {
                return "";
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
