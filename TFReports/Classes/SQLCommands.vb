Option Strict On
Option Explicit On
Imports System.Data.SqlClient
Imports Reports

Public Class SQLCommands
    'Dim mCnn As SqlConnection
    'Public Sub New(ByVal DBConnection As Reports.ReportsCollection.DBConnection)
    '    If DBConnection = Reports.ReportsCollection.DBConnection.TravelForce Then
    '        mCnn = New SqlConnection(NextDB.PanasoftConnections.TravelForcePanasoft)
    '        mCnn.Open()
    '    ElseIf DBConnection = ReportsCollection.DBConnection.TravelForceSQL18 Then
    '        mCnn = New SqlConnection(NextDB.DBConnections.TravelForce)
    '        mCnn.Open()
    '    Else
    '        mCnn = New SqlConnection(NextDB.PanasoftConnections.TravelForce)
    '        mCnn.Open()
    '    End If
    'End Sub
    'Public Function E00_Euronav(ByRef mReport As Reports.ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@TagId", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E00_Euronav"

    '    Return sqlComm

    'End Function
    'Public Function E02_BSPMonthReportByAirline(ByVal Domestic As Boolean, ByVal BSPYear As Integer, ByVal BSPMonth As Integer) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@Domestic", SqlDbType.Bit).Value = Domestic
    '    sqlComm.Parameters.Add("@BSPYear", SqlDbType.Int).Value = BSPYear
    '    sqlComm.Parameters.Add("@BSPMonth", SqlDbType.Int).Value = BSPMonth
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E02_BSPMonthReportByAirline"

    '    Return sqlComm

    'End Function
    'Public Function E03_BSPFortnightReportByTicket(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@BSPDate", SqlDbType.Date).Value = mReport.BSPFortDate
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E03_BSPFortnightReportByTicket"

    '    Return sqlComm
    'End Function
    'Public Function E04_TicketInfo(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand

    '    Dim pTickList As String = ""
    '    For i As Integer = 0 To mReport.TextEntryItemsCount
    '        If mReport.TextEntryItems(i).Length = 10 Then
    '            If pTickList.Length > 0 Then
    '                pTickList &= ","
    '            End If
    '            pTickList &= "'" & mReport.TextEntryItems(i) & "'"
    '        End If
    '    Next

    '    If pTickList.Length > 0 Then
    '        sqlComm.CommandType = CommandType.StoredProcedure
    '        sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '        sqlComm.Parameters.Add("@Docs", SqlDbType.NVarChar).Value = pTickList
    '        sqlComm.CommandText = "ATPIData.dbo.TFReports_E04_TicketInfo"
    '        Return sqlComm
    '    Else
    '        Throw New Exception("No tickets selected")
    '    End If

    'End Function
    'Public Function E05_ClientTurnover(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandTimeout = 180
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E05_ClientTurnover"
    '    Return sqlComm

    'End Function
    'Public Function E06_ProfitPerClientWithBudgetComparison(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '    sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '    sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '    sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '    sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '    sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '    sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '    sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '    sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E06_ProfitPerClientWithBudget"

    '    Return sqlComm

    'End Function
    'Public Function E07_ProfitPerOPSGroup(ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E07_ProfitPerOPSGroup"

    '    Return sqlComm
    'End Function
    'Public Function E08_ProfitPerClientGroup(ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E08_ProfitPerClientGroup"

    '    Return sqlComm
    'End Function

    'Public Function E09_ProfitPerOPSGroupWithExtra(ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E09_ProfitPerOPSGroupWithExtra"

    '    Return sqlComm
    'End Function
    'Public Function E10_ProfitPerClientGroupWithExtra(ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E10_ProfitPerClientGroupWithExtra"
    '    Return sqlComm
    'End Function
    'Public Function E11_ProfitPerOPSGroupWithPY(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal FromYTD As Date, ByVal ToYTD As Date, ByVal FromPY As Date, ByVal ToPY As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = DateFrom
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = DateTo
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = FromYTD
    '    sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = ToYTD
    '    sqlComm.Parameters.Add("@FromPY", SqlDbType.Date).Value = FromPY
    '    sqlComm.Parameters.Add("@ToPY", SqlDbType.Date).Value = ToPY
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E11_ProfitPerOPSGroupWithPY"


    '    Return sqlComm
    'End Function
    'Public Function E12_ProfitPerOPSGroupWithBudgetComparison(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '    sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '    sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '    sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '    sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '    sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '    sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '    sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '    sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E12_ProfitPerOPSGroupWithBudgetComparison"

    '    Return sqlComm

    'End Function
    'Public Function E40_ProfitPerAgentWithBudgetComparison(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '    sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '    sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '    sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '    sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '    sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '    sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '    sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '    sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E40_ProfitPerAgentwithBudgetComparison"
    '    Return sqlComm

    'End Function
    'Public Function E13_TicketAnalysis(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure

    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromIssue", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToIssue", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromDep", SqlDbType.Date).Value = mReport.E12_FromYTD
    '    sqlComm.Parameters.Add("@ToDep", SqlDbType.Date).Value = mReport.E12_ToYTD
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@TagId", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@UninvoicedOnly", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E13_TicketAnalysis"

    '    Return sqlComm
    'End Function
    'Public Function E16_DailyProfitReport(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E16_DailyProfitReport"
    '    Return sqlComm

    'End Function
    'Public Function E15_DailyProfitReportWithoutRINVA(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E15_DailyProfitReport"
    '    Return sqlComm

    'End Function
    'Public Function E17_ServiceFeeAnalysis(ByRef mReport As Reports.ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E17_ServiceFeeAnalysis"
    '    Return sqlComm

    'End Function
    'Public Function E18_AirTicketSales(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E18_AirTicketSales_1"
    '    Return sqlComm

    'End Function

    'Public Function E19_DailyProfitReportInvoicesWithTicketNumber(ByRef mReport As Reports.ReportsCollection, ByVal IW10ForAllCLients As Boolean) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@WithTicket", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    If mReport.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '    ElseIf mReport.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '    Else
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    End If
    '    sqlComm.Parameters.Add("@IW10AllClients", SqlDbType.Bit).Value = IW10ForAllCLients

    '    sqlComm.CommandTimeout = 120
    '    'changed stored procedure to cater for RINVA and accessing their valus from accounting transactions rather than commercial transactions
    '    'sqlComm.CommandText = "ATPIData.dbo.TFReports_E19_ProfitReportInvoicesWithIW"
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E19b_ProfitReportInvoicesWithIW"
    '    Return sqlComm

    'End Function
    'Public Function E19a_ProfitReportInvoicesTotals(ByRef mReport As Reports.ReportsCollection, ByVal IW10ForAllCLients As Boolean) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    If mReport.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '    ElseIf mReport.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '    Else
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    End If

    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E19a_ProfitReportInvoicesTotals"
    '    Return sqlComm

    'End Function
    'Public Function E20_HellasConfidence(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromIssueDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToIssueDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@IssueDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromInvoiceDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToInvoiceDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@InvoiceDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E20_HellasConfidence"
    '    Return sqlComm

    'End Function
    'Public Function E21_ReportByVerifiedUser(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.Parameters.Add("@FromVerifyDate", SqlDbType.DateTime).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToVerifyDate", SqlDbType.DateTime).Value = DateAdd(DateInterval.Hour, 24, mReport.Date1To)
    '    sqlComm.Parameters.Add("@VerifiedUserName", SqlDbType.NVarChar, 254).Value = mReport.GroupList
    '    sqlComm.CommandType = CommandType.Text
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.E21_ReportByVerifiedUser"

    '    Return sqlComm
    'End Function
    'Public Function E22_Euronav(ByRef mReport As Reports.ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E22_Euronav"

    '    Return sqlComm

    'End Function
    'Public Function E23_SeaChefs(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E23_SeaChefs"

    '    Return sqlComm

    'End Function
    'Public Function E23_SeaChefsX(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E23_SeaChefsX"
    '    Return sqlComm

    'End Function

    'Public Function E29_SeaChefsDetailed(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E29_SeaChefsDetailed"
    '    Return sqlComm

    'End Function
    'Public Function E24_ProfitPerAgentTotals(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E24_ProfitPerAgentTotals"

    '    Return sqlComm

    'End Function
    'Public Function E25_ProfitPerAgentTransactions(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E25_ProfitPerAgentTransactions"

    '    Return sqlComm

    'End Function
    'Public Function E28_OptimizationSavings(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E28_Optimization_Actions"
    '    Return sqlComm
    'End Function
    'Public Function E30_AirTicketsFullDetails(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@DateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked

    '    sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E30_AirTicketsFullDetails"
    '    Return sqlComm

    'End Function
    'Public Function E31_SeaChefsStatementCheck(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E31_SeaChefsStatementCheck"
    '    Return sqlComm

    'End Function
    'Public Function E36_SeaChefs_AllUnits(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E36_SeaChefs_AllUnits"
    '    Return sqlComm

    'End Function
    'Public Function E41_DailyProfitReportWithRINVAAnalysis(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E41_DailyProfitReportWithRINVAAnalysis"
    '    Return sqlComm
    'End Function
    'Public Function E42_AirTicketsWithFC(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To

    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E42_AirTicketsWithFC"
    '    Return sqlComm

    'End Function
    'Public Function E43_DailyProfitReportWithProvisionalAnalysis(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.CommandTimeout = 400
    '    'changed stored procedure to cater for RINVA and accessing their valus from accounting transactions rather than commercial transactions
    '    'sqlComm.CommandText = "ATPIData.dbo.TFReports_E43_DailyProfitReportWithProvisionalAnalysis"
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E43a_DailyProfitReportWithProvisionalAnalysis"
    '    Return sqlComm
    'End Function
    'Public Function E45_AirTicketSalesAll(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E45b_AirTicketSales_AllFields"
    '    Return sqlComm

    'End Function
    'Public Function E47_DailyProfitReportTotalsOnly(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E43d_DailyProfitReportWithProvisionalAnalysis"
    '    Return sqlComm
    'End Function
    'Public Function E48_DailyProfitReportTotalsPerInvoice(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E48_DailyProfitReportTotalsPerInvoice"
    '    Return sqlComm
    'End Function
    'Public Function E49_Optimization_Monthly_Report(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E49_Optimization_Monthly_Report"
    '    Return sqlComm
    'End Function
    'Public Function E50_Optimization_Annual_Report_by_Month(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@Year", SqlDbType.Int).Value = Conversion.Int(mReport.BSPMonthDate)
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "AmadeusReports.dbo.TFReports_E50_Optimization_Annual_Report_by_Month"
    '    Return sqlComm
    'End Function
    'Public Function E51_Daily_Profit_Totals_per_Category(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@NoOmit", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.CommandTimeout = 400
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E51_DailyProfitTotalsPerCategory"
    '    Return sqlComm
    'End Function
    'Public Function E52_Gaslog_Monthly_Statement(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E52_Gaslog_Monthly_Statement"
    '    Return sqlComm
    'End Function
    'Public Function E53_SeaChefs_InvoiceByDepartureDate(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '    sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '    sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '    sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '    sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E53_Sea_Chefs_Invoices_by_Departure_Date"
    '    Return sqlComm

    'End Function
    'Public Function E54_Client_Statement(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    'sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@IncludeOmit", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E54_Client_Statement"
    '    Return sqlComm

    'End Function
    'Public Function E55_Safety_Statement(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '    sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '    sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '    sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '    'sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '    sqlComm.Parameters.Add("@IncludeOmit", SqlDbType.Bit).Value = mReport.BooleanOption1
    '    sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '    sqlComm.CommandTimeout = 200
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E55_Safety_Statement"
    '    Return sqlComm

    'End Function
    'Public Function E56_Clients(ByRef mReport As Reports.ReportsCollection) As SqlCommand
    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_E56_ClientGroups"
    '    Return sqlComm

    'End Function

    'Public Function ClientList() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_ClientList"

    '    Return sqlComm

    'End Function
    'Public Function ClientGroupsAll() As SqlCommand

    '    ClientGroupsAll = New SqlCommand
    '    ClientGroupsAll = mCnn.CreateCommand
    '    ClientGroupsAll.CommandType = CommandType.StoredProcedure
    '    ClientGroupsAll.CommandText = "ATPIData.dbo.TFReports_ClientGroupsAll"

    'End Function
    'Public Function ClientGroupsSeaChefs() As SqlCommand

    '    ClientGroupsSeaChefs = New SqlCommand
    '    ClientGroupsSeaChefs = mCnn.CreateCommand
    '    ClientGroupsSeaChefs.CommandType = CommandType.StoredProcedure
    '    ClientGroupsSeaChefs.CommandText = "ATPIData.dbo.TFReports_ClientGroupsSeaChefs"
    'End Function
    'Public Function BSPMonths() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_BSPMonths"
    '    Return sqlComm

    'End Function
    'Public Function VerificationYears() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "AmadeusReports.dbo.TFReports_VerificationYears"
    '    Return sqlComm

    'End Function

    'Public Function TransactionYears() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_TransactionYears"
    '    Return sqlComm

    'End Function

    'Public Function BSPForthnights() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_BSPFortnights"
    '    Return sqlComm

    'End Function
    'Public Function AgentGroups() As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.TFReports_AgentGroups"
    '    Return sqlComm
    'End Function
    'Public Function Reader(cmm As SqlCommand) As SqlDataReader

    '    Reader = cmm.ExecuteReader

    'End Function

    'Public Function E12_ProfitPerOPSGroupWithBudgetComparisonX(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '    Dim sqlComm As New SqlCommand
    '    sqlComm = mCnn.CreateCommand
    '    sqlComm.Parameters.Add("@UserName", SqlDbType.NVarChar).Value = System.Environment.UserName
    '    sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '    sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '    sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '    sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '    sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '    sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '    sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '    sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '    sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '    sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '    sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '    sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '    sqlComm.CommandTimeout = 120
    '    sqlComm.CommandType = CommandType.StoredProcedure
    '    sqlComm.CommandText = "ATPIData.dbo.E12_ProfitPerOPSGroupWithBudgetComparisonX"
    '    Return sqlComm

    'End Function
End Class
