using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Reports;
using System.Data;

namespace TFReportsSQL
{
    public class StoredProcedures
    {
        readonly SqlConnection mCnn;
        readonly ReportsCollection mReport;
        public StoredProcedures(ReportsCollection.DBConnection cnn, ReportsCollection report)
        {
            mReport = report;
            if (cnn == ReportsCollection.DBConnection.TravelForce)
            {
                mCnn = new SqlConnection(NextDB.PanasoftConnections.TravelForcePanasoft);
                mCnn.Open();
            }
            else if (cnn == ReportsCollection.DBConnection.TravelForceSQL18)
            {
                mCnn = new SqlConnection(NextDB.DBConnections.TravelForce);
                mCnn.Open();
            }
            else
            {
                mCnn = new SqlConnection(NextDB.DBConnections.TravelForce);
                mCnn.Open();
            }
        }
        public SqlCommand E00_Euronav()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@TagId", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E00_Euronav";

            return sqlComm;
        }
        public SqlCommand E02_BSPMonthReportByAirline()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@Domestic", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.Parameters.Add("@BSPYear", SqlDbType.Int).Value = Convert.ToInt32(mReport.BSPMonthDate.Substring(0, 4));
            sqlComm.Parameters.Add("@BSPMonth", SqlDbType.Int).Value = Convert.ToInt32(mReport.BSPMonthDate.Substring(4, 2));
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E02_BSPMonthReportByAirline";

            return sqlComm;

        }
        public SqlCommand E03_BSPFortnightReportByTicket()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@BSPDate", SqlDbType.Date).Value = mReport.BSPFortDate;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E03_BSPFortnightReportByTicket";

            return sqlComm;
        }
        public SqlCommand E04_TicketInfo()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();

            string pTickList = "";
            for (int i = 0; i <= mReport.TextEntryItemsCount; i++)
            {
                if (mReport.TextEntryItems(i).Length == 10)
                {
                    if (pTickList.Length > 0)
                    {
                        pTickList += ",";
                    }
                    pTickList += "'" + mReport.TextEntryItems(i) + "'";
                }
            }

            if (pTickList.Length > 0)
            {
                sqlComm.CommandType = CommandType.StoredProcedure;
                sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
                sqlComm.Parameters.Add("@Docs", SqlDbType.NVarChar).Value = pTickList;
                sqlComm.CommandText = "ATPIData.dbo.TFReports_E04_TicketInfo";
                return sqlComm;
            }
            else
            {
                throw new Exception("No tickets selected");
            }

        }
        public SqlCommand E05_ClientTurnover()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 180;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E05_ClientTurnover";
            return sqlComm;

        }
        public SqlCommand E06_ProfitPerClientWithBudgetComparison()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD;
            sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD;
            sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD;
            sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD;
            sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr;
            sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr;
            sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear;
            sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1;
            sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E06_ProfitPerClientWithBudget";

            return sqlComm;

        }
        public SqlCommand E07_ProfitPerOPSGroup()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E07_ProfitPerOPSGroup";

            return sqlComm;
        }


        public SqlCommand E08_ProfitPerClientGroup()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E08_ProfitPerClientGroup";

            return sqlComm;
        }
        public SqlCommand E09_ProfitPerOPSGroupWithExtra()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E09_ProfitPerOPSGroupWithExtra";

            return sqlComm;
        }
        public SqlCommand E10_ProfitPerClientGroupWithExtra()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E10_ProfitPerClientGroupWithExtra";
            return sqlComm;
        }


        public SqlCommand E11_ProfitPerOPSGroupWithPY()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.FromYTD;
            sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.ToYTD;
            sqlComm.Parameters.Add("@FromPY", SqlDbType.Date).Value = mReport.FromPYTD;
            sqlComm.Parameters.Add("@ToPY", SqlDbType.Date).Value = mReport.ToPYTD;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E11_ProfitPerOPSGroupWithPY";


            return sqlComm;
        }
        public SqlCommand E40_ProfitPerAgentWithBudgetComparison(int TagGroup)
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD;
            sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD;
            sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD;
            sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD;
            sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr;
            sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr;
            sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear;
            sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1;
            sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E40_ProfitPerAgentwithBudgetComparison";
            return sqlComm;

        }

        public SqlCommand E13_TicketAnalysis()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;

            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromIssue", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToIssue", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromDep", SqlDbType.Date).Value = mReport.E12_FromYTD;
            sqlComm.Parameters.Add("@ToDep", SqlDbType.Date).Value = mReport.E12_ToYTD;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@TagId", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@UninvoicedOnly", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E13_TicketAnalysis";

            return sqlComm;
        }
        public SqlCommand E16_DailyProfitReport()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E16_DailyProfitReport";
            return sqlComm;

        }
        public SqlCommand E15_DailyProfitReportWithoutRINVA()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E15_DailyProfitReport";
            return sqlComm;

        }
        public SqlCommand E17_ServiceFeeAnalysis()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E17_ServiceFeeAnalysis";
            return sqlComm;

        }
        public SqlCommand E18_AirTicketSales()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E18_AirTicketSales_1";
            return sqlComm;

        }

        public SqlCommand E19_DailyProfitReportInvoicesWithTicketNumber(bool IW10ForAllCLients)
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@WithTicket", SqlDbType.Bit).Value = mReport.BooleanOption1;
            if (mReport.ByClient == Reports.ReportsCollection.ClientReportType.AllClients)
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = "";
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0;
            }
            else if (mReport.ByClient == Reports.ReportsCollection.ClientReportType.ByClient)
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0;
            }
            else
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = "";
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            }
            sqlComm.Parameters.Add("@IW10AllClients", SqlDbType.Bit).Value = IW10ForAllCLients;

            sqlComm.CommandTimeout = 120;
            //changed stored procedure to cater for RINVA and accessing their valus from accounting transactions rather than commercial transactions
            //sqlComm.CommandText = "ATPIData.dbo.TFReports_E19_ProfitReportInvoicesWithIW"
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E19b_ProfitReportInvoicesWithIW";
            return sqlComm;

        }
        public SqlCommand E19a_ProfitReportInvoicesTotals()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            if (mReport.ByClient == Reports.ReportsCollection.ClientReportType.AllClients)
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = "";
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0;
            }
            else if (mReport.ByClient == Reports.ReportsCollection.ClientReportType.ByClient)
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0;
            }
            else
            {
                sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = "";
                sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            }

            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E19a_ProfitReportInvoicesTotals";
            return sqlComm;

        }
        public SqlCommand E20_HellasConfidence()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromIssueDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToIssueDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@IssueDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromInvoiceDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToInvoiceDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@InvoiceDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E20_HellasConfidence";
            return sqlComm;

        }
        public SqlCommand E21_ReportByVerifiedUser()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.Parameters.Add("@FromVerifyDate", SqlDbType.DateTime).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToVerifyDate", SqlDbType.DateTime).Value = mReport.Date1To.AddHours(24);
            sqlComm.Parameters.Add("@VerifiedUserName", SqlDbType.NVarChar, 254).Value = mReport.GroupList;
            sqlComm.CommandType = CommandType.Text;
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.E21_ReportByVerifiedUser";

            return sqlComm;
        }
        public SqlCommand E22_Euronav() // ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E22_Euronav";

            return sqlComm;

        }
        public SqlCommand E23_SeaChefsX()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E23_SeaChefsX";
            return sqlComm;

        }

        public SqlCommand E29_SeaChefsDetailed()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E29_SeaChefsDetailed";
            return sqlComm;

        }
        public SqlCommand E24_ProfitPerAgentTotals()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E24_ProfitPerAgentTotals";

            return sqlComm;

        }
        public SqlCommand E25_ProfitPerAgentTransactions()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E25_ProfitPerAgentTransactions";

            return sqlComm;

        }
        public SqlCommand E28_OptimizationSavings()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E28_Optimization_Actions";
            return sqlComm;
        }
        public SqlCommand E30_AirTicketsFullDetails()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@DateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;

            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E30a_AirTicketsFullDetails";
            return sqlComm;

        }
        public SqlCommand E31_SeaChefsStatementCheck()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E31_SeaChefsStatementCheck";
            return sqlComm;

        }
        public SqlCommand E36_SeaChefs_AllUnits()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E36_SeaChefs_AllUnits";
            return sqlComm;

        }
        public SqlCommand E41_DailyProfitReportWithRINVAAnalysis()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E41_DailyProfitReportWithRINVAAnalysis";
            return sqlComm;
        }
        public SqlCommand E42_AirTicketsWithFC()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;

            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E42_AirTicketsWithFC";
            return sqlComm;

        }

        public SqlCommand E43_DailyProfitReportWithProvisionalAnalysis()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.Parameters.Add("@CheckUninvDepDate", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.Parameters.Add("@UninvoicedLimit", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.CommandTimeout = 400;
            //changed stored procedure to cater for RINVA and accessing their valus from accounting transactions rather than commercial transactions
            //sqlComm.CommandText = "ATPIData.dbo.TFReports_E43_DailyProfitReportWithProvisionalAnalysis"
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E43a_DailyProfitReportWithUninvoicedLimit";
            //sqlComm.CommandText = "ATPIData.dbo.TFReports_E43a_DailyProfitReportWithProvisionalAnalysis";
            return sqlComm;
        }
        public SqlCommand E45_AirTicketSalesAll()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E45b_AirTicketSales_AllFields";
            return sqlComm;

        }
        public SqlCommand E47_DailyProfitReportTotalsOnly()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E43d_DailyProfitReportWithProvisionalAnalysis";
            return sqlComm;
        }
        public SqlCommand E48_DailyProfitReportTotalsPerInvoice()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E48_DailyProfitReportTotalsPerInvoice";
            return sqlComm;
        }
        public SqlCommand E49_Optimization_Monthly_Report()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E49_Optimization_Monthly_Report";
            return sqlComm;
        }
        public SqlCommand E50_Optimization_Annual_Report_by_Month()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@Year", SqlDbType.Int).Value = Convert.ToInt32(Math.Floor(Convert.ToDouble(mReport.BSPMonthDate)));
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E50_Optimization_Annual_Report_by_Month";
            return sqlComm;
        }
        public SqlCommand E51_Daily_Profit_Totals_per_Category()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = new DateTime(mReport.Date1From.Year, 1, 1);
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.CommandTimeout = 400;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E51_DailyProfitTotalsPerCategory";
            return sqlComm;
        }
        public SqlCommand E52_Gaslog_Monthly_Statement()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E52_Gaslog_Monthly_Statement";
            return sqlComm;
        }
        public SqlCommand E53_SeaChefs_InvoiceByDepartureDate()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked;
            sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From;
            sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To;
            sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E53_Sea_Chefs_Invoices_by_Departure_Date";
            return sqlComm;

        }
        public SqlCommand E54_Client_Statement()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            //sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@IncludeOmit", SqlDbType.Bit).Value = mReport.BooleanOption1;
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E54_Client_Statement";
            return sqlComm;

        }
        public SqlCommand E55_Safety_Statement()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            //sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
            //sqlComm.Parameters.Add("@IncludeOmit", SqlDbType.Bit).Value =  mReport.BooleanOption1;
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E55_Safety_Statement";
            return sqlComm;

        }
        public SqlCommand E56_Clients()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E56_ClientGroups";
            return sqlComm;

        }
        public SqlCommand E57_TUI_030366()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E57_TUI_030366";
            return sqlComm;

        }
         public SqlCommand E59_AirTicketsByInvoiceDate()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E59_AirTicketsByInvoiceDate";
            return sqlComm;

        }
        public SqlCommand E63_AirTicketSalesTemenos()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E63_AirTicketsTemenos";
            return sqlComm;

        }
        public SqlCommand E64_LowestClasses()
        { 
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E64_LowestClasses";
            return sqlComm;

        }
        public SqlCommand E65_OpsSales()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E65_OpsSales";
            return sqlComm;

        }
        public SqlCommand E66_PurchasesPerAirline()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@FromDateCY", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDateCY", SqlDbType.Date).Value = mReport.Date1To;
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E66_TicketPurchasesByDepartureDate";
            return sqlComm;

        }
        public SqlCommand E67_Columbia()
        {
            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID;
            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer;
            sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From;
            sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To;
            //sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet; // 0 = All 1 = Uninvoiced 2 = Invoiced
            sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace("\r\n", "|");
            sqlComm.CommandTimeout = 200;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E67_Columbia";
            return sqlComm;

        }
        public SqlCommand ClientList()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_ClientList";

            return sqlComm;

        }
        public SqlCommand ClientGroupsAll()
        {
            SqlCommand tempClientGroupsAll = mCnn.CreateCommand();
            tempClientGroupsAll.CommandType = CommandType.StoredProcedure;
            tempClientGroupsAll.CommandText = "ATPIData.dbo.TFReports_ClientGroupsAll";

            return tempClientGroupsAll;
        }
        public SqlCommand ClientGroupsSeaChefs()
        {
            SqlCommand tempClientGroupsSeaChefs = mCnn.CreateCommand();
            tempClientGroupsSeaChefs.CommandType = CommandType.StoredProcedure;
            tempClientGroupsSeaChefs.CommandText = "ATPIData.dbo.TFReports_ClientGroupsSeaChefs";
            return tempClientGroupsSeaChefs;
        }
        public SqlCommand BSPMonths()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_BSPMonths";
            return sqlComm;

        }
        public SqlCommand VerificationYears()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "AmadeusReports.dbo.TFReports_VerificationYears";
            return sqlComm;

        }

        public SqlCommand TransactionYears()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_TransactionYears";
            return sqlComm;

        }

        public SqlCommand BSPForthnights()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_BSPFortnights";
            return sqlComm;

        }
        public SqlCommand AgentGroups()
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_AgentGroups";
            return sqlComm;
        }
        public SqlDataReader Reader(SqlCommand cmm)
        {

            return cmm.ExecuteReader();

        }

        public SqlCommand E12_ProfitPerOPSGroupWithBudgetComparisonX(int TagGroup)
        {

            SqlCommand sqlComm = mCnn.CreateCommand();
            sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName;
            sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup;
            sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr;
            sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr;
            sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD;
            sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD;
            sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD;
            sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD;
            sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr;
            sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr;
            sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear;
            sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1;
            sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth;
            sqlComm.CommandTimeout = 120;
            sqlComm.CommandType = CommandType.StoredProcedure;
            sqlComm.CommandText = "ATPIData.dbo.TFReports_E12_ProfitPerOPSGroupWithBudgetComparisonX"; // "ATPIData.dbo.E12_ProfitPerOPSGroupWithBudgetComparisonX";
            return sqlComm;

        }

    }
}
