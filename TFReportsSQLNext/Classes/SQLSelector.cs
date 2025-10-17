using Microsoft.Data.SqlClient;
using PNRHistoryNext;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFReportsSQLNext.Classes
{
    public class SQLSelector
    {

        PNRHistoryNext.PNRs pPNR = new();
        GaslogReports.E33_Collection pE33_Gaslog;
        GaslogReports.E35_Collection pE35_Gaslog;
        GDSImport.GDSImportItems pGDSImport = new();
        ReportsNext.ReportsItem mobjSelectedReport;
        ReportsNext.ReportsCollection mobjReports;
        SqlDataReader mSQLDataReader;
        public string ReportTitle { get; set; } = "";
        public string ReportFileName { get; set; } = "";
        public SQLSelector(ReportsNext.ReportsItem selectedreport, ReportsNext.ReportsCollection reports)
        {
            mobjReports = reports;
            mobjSelectedReport = selectedreport;
        }
        public void PrepareDisjointedData()
        {
            if (mobjSelectedReport == null || mobjReports == null)
            {
                throw new Exception("No report selected");
            }
            SqlStatement(false);

        }
        public System.Data.DataSet PrepareDataSet()
        {
            if (mobjSelectedReport == null || mobjReports == null)
            {
                throw new Exception("No report selected");
            }
            PrepareFileName();
            Microsoft.Data.SqlClient.SqlDataAdapter dscmd = new(SqlStatement(false));
            System.Data.DataSet ds = new();
            dscmd.Fill(ds);
            if (ds.Tables.Count == 0)
            {
                throw new Exception("No tables returned");
            }
            return ds;
        }
        void PrepareFileName()
        {
            ReportTitle = mobjSelectedReport.ReportName.Replace(" ", "_");
            if (mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByClient)
            {
                ReportTitle += "_" + mobjReports.SelectedCustomer;
            }
            else if (mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByGroup)
            {
                ReportTitle += "_" + mobjReports.CustomerGroup;
            }
            if (mobjSelectedReport.Date1Text.Length > 0)
            {
                ReportTitle += "_" + mobjReports.Date1From.ToString("yyyyMMdd");
            }
            if (mobjSelectedReport.Date2Text.Length > 0)
            {
                ReportTitle += "-" + mobjReports.Date1To.ToString("yyyyMMdd");
            }

            ReportFileName += $"{ReportTitle}.xlsx";
        }
        public Microsoft.Data.SqlClient.SqlCommand SqlStatement(bool withReader)
        {
            Microsoft.Data.SqlClient.SqlCommand sqlCmd = new();
            StoredProcedures sql = new(mobjSelectedReport.DBConnection, mobjReports);
            switch (mobjSelectedReport.Index)
            {               
                case 7:
                    sqlCmd = sql.E07_ProfitPerOPSGroup();
                    break;
                case 8:
                    sqlCmd = sql.E08_ProfitPerClientGroup();
                    break;      
                case 10:
                    sqlCmd = sql.E10_ProfitPerClientGroupWithExtra();
                    break;
                case 12:
                    sqlCmd = sql.E12_ProfitPerOPSGroupWithBudgetComparisonX(149);
                    break;
                case 18:
                    sqlCmd = sql.E18_AirTicketSales();
                    break;
                case 19:
                    sqlCmd = sql.E19_DailyProfitReportInvoicesWithTicketNumber(false);
                    break;
                case 20:
                    sqlCmd = sql.E20_HellasConfidence();
                    break;
                case 22:
                    sqlCmd = sql.E22_Euronav();
                    break;
                case 23:
                    sqlCmd = sql.E23_SeaChefsX();
                    break;
                case 24:
                    sqlCmd = sql.E24_ProfitPerAgentTotals();
                    break;
                case 25:
                    sqlCmd = sql.E25_ProfitPerAgentTransactions();
                    break;
                case 26:
                    break;
                case 27:
                    break;
                case 29:
                    sqlCmd = sql.E29_SeaChefsDetailed();
                    break;
                case 30:
                    sqlCmd = sql.E30_AirTicketsFullDetails();
                    break;
                case 32:
                    pPNR.ReadPNRs(mobjReports.Date1From, mobjReports.Date1To);
                    break;
                case 33:
                    pE33_Gaslog = new GaslogReports.E33_Collection(mobjReports.Date1From, mobjReports.Date1To, "012212");
                    break;
                case 34:
                    break;
                case 35:
                    pE35_Gaslog = new GaslogReports.E35_Collection(mobjReports.Date1From, mobjReports.Date1To, "012217");
                    break;
                case 36:
                    sqlCmd = sql.E36_SeaChefs_AllUnits();
                    break;
                case 37:
                    sqlCmd = sql.E19_DailyProfitReportInvoicesWithTicketNumber(true);
                    break;
                case 38:
                    sqlCmd = sql.E18_AirTicketSales();
                    break;
                case 39:
                    pGDSImport.Read();
                    break;                           
                case 43:
                    sqlCmd = sql.E43_DailyProfitReportWithProvisionalAnalysis();
                    break;
                case 44:
                    break;
                case 45:
                    sqlCmd = sql.E45_AirTicketSalesAll();
                    break;
                case 46:
                    sqlCmd = sql.E19a_ProfitReportInvoicesTotals();
                    break;
                case 47:
                    sqlCmd = sql.E47_DailyProfitReportTotalsOnly();
                    break;
                case 48:
                    sqlCmd = sql.E48_DailyProfitReportTotalsPerInvoice();
                    break;
                case 51:
                    sqlCmd = sql.E51_Daily_Profit_Totals_per_Category();
                    break;
                case 53:
                    sqlCmd = sql.E53_SeaChefs_InvoiceByDepartureDate();
                    break;
                case 54:
                    sqlCmd = sql.E54_Client_Statement();
                    break;
                case 56:
                    sqlCmd = sql.E56_Clients();
                    break;
                case 57:
                    sqlCmd = sql.E57_TUI_030366();
                    break;
                case 58:
                    break;
                case 59:
                    sqlCmd = sql.E59_AirTicketsByInvoiceDate();
                    break;
                case 60:
                    sqlCmd = sql.E55_Safety_Statement();
                    break;
                case 61:
                    pE33_Gaslog = new GaslogReports.E33_Collection(mobjReports.Date1From, mobjReports.Date1To, "012629");
                    break;
                case 62:
                    pE35_Gaslog = new GaslogReports.E35_Collection(mobjReports.Date1From, mobjReports.Date1To, "012629");
                    break;
                case 63:
                    sqlCmd = sql.E63_AirTicketSalesTemenos();
                    break;
                case 64:
                    sqlCmd = sql.E64_LowestClasses();
                    break;
                case 65:
                    sqlCmd = sql.E65_OpsSales();
                    break;
                case 66:
                    sqlCmd = sql.E66_PurchasesPerAirline();
                    break;
                case 67:
                    sqlCmd = sql.E67_Columbia();
                    break;
case 68:
                    sqlCmd = sql.E68_GDSImportedPendingItems();
                    break;
                default:
                    throw new Exception("Report not implemented");
            }
            if (withReader)
            {
                mSQLDataReader = sql.Reader(sqlCmd);
            }

            // Add parameters as needed
            return sqlCmd;
        }
    }
}
