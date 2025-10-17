Imports System.Data.SqlClient
Imports System.Threading

Public Class FrmTFReports
    Private Const VersionText As String = "V 2025.15 17/09/2025"
    ' 11/02/2019 (TFReports_2_0_0_32)
    ' 1. Report 18 Air Ticket Sales
    '    Non Commercial Transactions RINVA, Etc do not have entry in CommercialTransactionCards and give NULL values that were not trapped in thq SQL query causing the program to crash
    ' 13/02/2019 (TFReports_2_0_0_33)
    ' 1. Report 12 Profit Per Client Group with Budget Comparison
    '    Add a new category 99 - OTHER to show budget for entries not connected to actual clients (new clients, expansion for existing clients, etc
    '    This required the addition of 2 new fields in the TFClientBudget table for the Code and Name that will be shown in the report
    ' 06/03/2019 (TFReports_2_0_0_34)
    ' 1. Report 15 Daily Profit Report
    '    Increase the timeout to 240
    ' 10/04/2019 (TFReports_2_0_0_36)
    ' 1. Added report 19 which is a copy of report 17 with extra columns for Ticket Number, Face Value, Taxes, Commission, Discount, Cancellation Fee and TF
    ' 12/04/2019 (TFReports_2_0_0_37)
    ' 1. Added a new option "All clients"
    ' 22/04/2019 (TFReports_2_0_0_38)
    ' 1. Added IW11 to reports 15 and 19
    ' 2. Added a new option to report 19 "With tickets" and removed report 17 which is the same as old report 19 but without tickets
    ' 24/04/2019 (TFReports_2_0_0_39)
    ' 1. Corrected an error in report 19 where the IW amount was not being multiplied by the currency rate
    ' 2. Added new report 17 Service Fee Analysis
    ' 03/05/2019 (TFReports_2_0_0_40)
    ' 1. Added Uninvoiced Pax and Provisional Profit in 15 Daily Profit Report
    ' 27/05/2019 (TFReports_2_0_0_41)
    ' 1. added report 20 Hellas Confidence. This is a cut-down version of report 18.
    ' 29/05/2019 (TFReports_2_0_0_42)
    ' 1. Report 18 was highlighting the omitted invoices incorrectly
    '    Spreadsheet.E19_DailyProfitReportInvoicesWithTicketNumber()
    '    the line        : If mdsDataSet.Tables(0).Rows(i).Item(60) Then
    '    was changed to  : If mdsDataSet.Tables(0).Rows(i).Item(63) Then
    ' 12/06/2019 (TFReports_2_0_0_43)
    '    Added a new option in the 2 From/To dates. The checkbox allows the user to ignore one of the two dates. If e ignores both, the options are not valid and the program will not run.
    '    Specifically in the case of report 20. Hellas Confidence, the user can select From/To Issue Date and/or From/To Invoice Date.
    '    The SQL for report 20 Hellas Confidence was changed to take into account these 2 options
    ' 20/06/2019 (TFReports_2_0_0_44)
    ' SQL Statement for report 18 crashed when selecting client group because SelectedCustomer is nothing
    ' 21/06/2019 (TFReports_2_0_0_45)
    ' Added a new column in report 18 - Ticketing Airline
    ' 24/06/2019 (TFReports_2_0_0_46)
    ' Added 3 new columns in report 18 - Salesman, Issuing Agent, Creator Agent
    ' 24/06/2019 (TFReports_2_0_0_47)
    ' Report 18 - All Customers option was not working and I also aded one more column with the routing
    ' Report 18 - Added a new option to select airlines
    ' 24/06/2019 (TFReports_2_0_0_47)
    ' Changed text box for Airline codes to CAPITALS
    ' after a file is created, the directory is opened
    ' 12/07/2019 (TFReports_2_0_0_50)
    ' Added Issuing Airline to Report 19
    ' Added New Group "Optimize Reports" and new report "21 Report by Verified User"
    ' This required the addition of new parameters
    ' 1 - to access the AmadeusReports database and 
    ' 2 - to reuse the listbox in the options to show the Agents
    ' 29/10/2019 (TFReports_2_0_0_53)
    ' Added field "Reference" to report 18
    ' 21/11/2019 (TFReports_2_0_0_54)
    ' Added fields "Departure Date" and "Arrival Date" to report 18
    ' 10/02/2020 (TFReports_2_0_0_55)
    ' Remove the check CommercialTransactions.CustomDescription3 <> '' from reports 15 and 19
    ' 17/02/2020 (TFReports_2_0_0_56)
    ' Add column Account Code in report 18
    ' 25/02/2020 (TFReports_2_0_0_57)
    ' Added 2 new columns to report 18. Connected Document and Passenger Remarks
    ' Created Report 22-Euronav which is similar to 00-Euronav but is based on report 18
    ' 22/05/2020 (TReports_2_0_0_58)
    ' Changed the DB Connection for the new Panasoft database
    ' 29/05/2020 (TReports_2_0_0_59)
    ' Report 22-Euronav - Change the SQL statement to check invoice date and not commercial transaction date
    ' 29/05/2020 (TReports_2_0_0_60)
    ' Report 22-Euronav - Change the SQL statement to ignore cancelled documents
    ' 01/06/2020 (TReports_2_0_0_61)
    ' Report 22-Euronav - Added WHERE to exclude cancellation documents and omitted documents
    ' V 2020.0 30/06/2020
    ' The new generation
    ' using port 453334 to connect to the SQL database
    ' V 2020.1 22/07/2020
    ' Added new column "Provisional Uninvoiced T/O" in report 15-Daily Profit
    ' V 2020.2 23/07/2020
    ' Some SQL queries were still referring to the server EUDC-SQLCL14. This has now been fixed
    ' V 2020.3 04/09/2020
    ' 15. Daily Report. Client groups had incorrect YTD Pax because it ignored client codes that were active during the year 
    '                   but inactive in the requested period
    ' added "OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0" in the SQL where client codes are selected
    ' V 2020.4 02/02/2021
    ' 23. Sea Chefs statement
    ' V 2020.5 02/02/2021
    ' Small error in Filename for report 23 fixed
    ' V 2020.6 10/02/2021
    ' Fixed the SQL statement for report 23. It was not checking the client number correctly.
    ' V 2020.7 10/02/2021
    ' Changed the form. Removed tool strip button and made it normal button because tool strip
    ' does not cause validation and confuses the user.
    ' V 2020.8 15/02/2021
    ' Changed SQL for 23 to cater for multiple records returned by subquery
    ' also added timeout=120 to avoid timeout problem
    ' V 2020.9 175/02/2021
    ' 23 Seachefs Statement. CHanged the Description Column to add Booked By and TRID
    ' Added 2 columns, client code and client name
    ' V 2021.1 09/03/2021
    ' Added reports 24 and 25 for Profit per agent
    ' V 2021.3 11/03/2021
    ' Changed the query for report 23 to check only for last segment departure date
    ' V 2021.4 23/03/2021
    ' Added departure date to report 22 - Euronav
    ' V 2021.5 07/05/2021
    ' Changed the calculation of Provisional Profit from uninvoiced Pax in report 15
    ' added extra checks for cancellation documents in report 23
    ' V 2021.6 11/05/2021
    ' changed the SQL query for 15 Daily Profit Report to separate air tickets from services
    ' V 2021.7 12/05/2021
    ' added report 28 - Optimization Savings
    ' V 2021.8 22/06/2021
    ' renamed report 23 to 29 Sea Chefs Detailed
    ' and redesigned report 23 to match Sea Chefs' latest requirements
    ' V 2021.9 01/07/2021
    ' fixed SQL 29 to avoid NULL references
    ' V 2021.10 08/07/2021
    ' fixed SQL 23 to avoid NULL references
    ' V 2021.11 12/07/2021
    ' fixed SQL 15 and 19 to prevent duplicate cost of hotels in some cases
    ' V 2021.12 23/07/2021
    ' change 23 for the layout changes requested by Sea Chefs
    ' also introduced a parameter to limit the selectable Customer Groups when running the Sea Chefs report
    ' V 2021.13 28/07/2021
    ' Changed 23 to cater for a different structure of the client groups and Sea Chefs Business Units
    ' V 2021.14 29/07/2021
    ' Added new BUs to 23 Sea Chefs
    ' V 2021.15 29/07/2021
    ' forgot two END statement in sql QUERY
    ' V 2021.16 29/07/2021
    ' Supplier number also changed in SQL Query for 23
    ' V 2021.17 18/08/2021
    ' Customer Group SQL Command had an infinite loop because of the brackets in "With Customer_Group()"
    ' I removed all With statements in SQLCommands.vb
    ' V 2021.19 13/09/2021
    ' Fixed 23 to not have any NULL values so that it can run for individual Client Codes
    ' V 2021.20 21/09/2021
    ' added more codes to SQL for 23
    ' V 2021.21 22/09/2021
    ' Added one more WHERE clause in SQL 23 to avoid accounting entries
    ' AND DocTypes.UsesFSD = 1
    ' V 2021.22 29/09/2021
    ' SQL 23 - added a new table for Sea Chefs Business Units
    ' V 2021.23 20/10/2021
    ' SQL 30 added
    ' V 2021.24 24/11/2021
    ' added cancelled doc analysis to report 18
    ' V 2021.25 27/12/2021
    ' Problem with null values in report 19
    ' V 2022.01 10/01/2022
    ' Changed report 23 to check for last departure date and not first departure date
    ' V 2022.02 19/01/2022
    ' Added report 31
    ' V 2022.03 02/03/2022
    ' Exclude Doc Type 109 in Report 15. Alos exclude clients from AmadeusReports.dbo.TFReportExclude
    ' V 2022.04 08/10/2022
    ' Removed DocType 134 (RINVA) from the excluded document types 15 Profit Report
    ' V 2022.05 29/12/2022
    ' Added Report 32 Accounting QC for Selling Price
    ' added new projects PNRHistory and Reports
    ' V 2022.06 27/1/2023
    ' added Other Services column to report 18
    ' V 2022.07 27/1/2023
    ' added Client Team column to report 18
    ' V 2022.08 30/1/2023
    ' Changed report 29 & 31 to select supplier name depending on client code
    ' created stored procedures:
    ' ATPIData.dbo.TFReports_E29_SeaChefsDetailed
    ' ATPIData.dbo.TFReports_E31_SeaChefsStatementCheck
    ' V 2022.09 14/2/2023
    ' added reports 33 & 35 for Gaslog
    ' V 2022.10 15/2/20230
    ' V 2022.11 15/2/20230
    ' added report 36 which is a clone of 23 and works for all Sea Chefs Business Units automatically
    ' V 2022.12 16/2/2023
    ' added client number and some other fields to report 36
    ' V 2023.2 6/3/2023
    ' fixed NextDB connection for report 32
    ' V 2023.3 9/3/2023
    ' added reports 37 and 38
    ' V 2023.4 9/3/2023
    ' added stored procedure for 19 and 37
    ' V 2023.5 10/3/2023
    ' missed a column in report 38. Now fixed
    ' V 2023.6 21/3/2023
    ' converting all SQL to stored procedures and adding log file
    ' V 2023.7 27/4/2023
    ' Added report 39 and added all client profile elements to report 30
    ' V 2023.8 27/4/2023
    ' Finish off converting to Stored Procedures and adding to log file
    ' V 2023.9 08/05/2023
    ' Revert report 12 to previous version
    ' V 2023.10 12/05/2023
    ' Revert Report 23 to old version
    ' Add Report 40
    ' V 2023.11 17/05/2023
    ' add new columns to 18 (sell, buy, profit, pax+-)
    ' V 2023.12 25/05/2023
    ' Fix RXRegex in PNRHistory.Remarks
    ' V 2023.13 01/06/2023
    ' Added new columns to Report 32 (Dummy/FixedMarkup/Void/Refund)
    ' V 2023.14 12/06/2023
    ' added report 41 Profit Report with RINVA Analysis
    ' V 2023.15 12/07/2023
    ' Added report 42 Air Ticket with FC
    ' Added FC Column to Report 19
    ' V 2023.16 20/07/2023
    ' Added new option to Report 23 - Separate Worksheet per Booked By
    ' V 2023.17 18/08/2023
    ' Fixed bug where date2 checkbox was reset when chkCheck was clicked
    ' V 2023.18 05/09/2023
    ' added report 43 Daily Profit with Provisional Analysis
    ' V 2023.20 15/09/2023
    ' added report 45
    ' V 2023.21 09/10/2023
    ' added fare basis to report 45
    ' V 2023.22 06/11/2023
    ' added report 19a
    ' added Client Group and Ops Team to report 19
    ' V 2023.23 20/11/2023
    ' Changed Stored Procedures for Report 43 to 43a and 19 to 19b
    ' V 2023.24 20/11/2023
    ' added cabin class and duration to report 45
    ' V 2023.25 08/12/2023
    ' added report 47 Profit Report Totals
    ' added Ops group selector to report 12
    ' V 2023.26 12/12/2023
    ' added IW5 columns to report 47
    ' V 2023.27 13/12/2023
    ' added details to the IW5 columns in report 47
    ' V 2023.28 20/12/2023
    ' added omit option to report 47
    ' added report 48
    ' V 2024.01 05/01/2024
    ' added reports 49 & 50
    ' V 2024.02 20/02/2024
    ' added report 51
    ' V 2024.03 29/03/2024
    ' fixed report 32
    ' if there was no reason for travel the program crashed
    ' V 2024.04 29/03/2024
    ' added report 53
    ' V 2024.05 04/04/2024
    ' added report 54
    ' V 2024.06 08/05/2024
    ' fixed optimization reports 49 & 50 to read from correct SQL server
    ' V 2024.07 12/06/2024
    ' add option ignore OMIT/VOID to report 18
    ' I want this for the lowest class of service report in TicketStatusNext
    ' V 2024.08 20/06/2024
    ' added report 56 - Clients per Group
    ' V 2024.09 26/06/2024
    ' Changed report 28 to show POSTPONE/REJECT reason
    ' V 2024.10 27/06/2024
    ' added Count of Passive and non-air to report 32
    ' V 2024.11 27/06/2024
    ' added a check for string length     
    ' switch ($"{CostCentre}0000".Substring(0, 4))
    ' in E35_Item.TransactionKey
    ' V 2024.12 17/07/2024
    ' created all the worksheets in report 28 - Optimization Savings
    ' V 2024.13 06/08/2024
    ' added analysis per hour to report 28 Action Totals sheet
    ' V 2024.14 08/08/2024
    ' add TOTAL row for report 28 sheet Action Totals
    ' V 2024.15 08/08/2024
    ' added Report parameter which sets initial values for the report
    ' V 2024.16 11/09/2024
    ' added report 57 TUI 030366
    ' V 2024.17 22/10/2024
    ' added column Net Remit to report 30
    ' V 2024.18 18/11/2024
    ' fixed Potential downsell in report 28. It was showing pax total and not PNR total
    ' V 2024.19 04/12/2024
    ' Reports.E28_ActionReasonItem -> public int Ranking(int index)
    ' added division by zero check
    ' V 2024.20 10/12/2024
    ' added Inactive to Report 12
    ' V 2025.01 24/01/2025
    ' added By Client sheet to 28
    ' V 2025.02 30/01/2025
    ' added With Fare Basis option to 55
    ' V 2025.03 30/01/2025
    ' added report 60
    ' V 2025.04 03/02/2025
    ' added booking date and other minor details to 60
    ' V 2025.05 21/02/2025
    ' fixed 28 tabs with group details were not sorting properly
    ' V 2025.06 19/03/2025
    ' added report 61 and 62
    ' V 2025.07 28/03/2025
    ' Added invoice currency option to report 22
    ' V 2025.08 02/05/2025
    ' Report 63 Temenos
    ' V 2025.09 15/05/2025
    ' Added Report 64
    ' V 2025.10 22/05/2025
    ' Added Report 66
    ' V 2025.11 22/05/2025
    ' Changed Report 66 to use from/to date rather than specific month
    ' V 2025.12 17/06/2025
    ' Changed report 63 to use Invoice Issue Date instead of Commercial Transaction Date
    ' V 2025.13 07/07/2025
    ' Report 66 - There was a problem calculating the Indeces
    ' V 2025.14 05/09/2025
    ' Report 38 - added 2 new columns Supplier Code and Supplier Name
    ' V 2025.15 17/09/2025
    ' added report 67 Columbia for Bellina
    Private ds As New DataSet

    Private WithEvents MySpreadSheet As SpreadSheet
    Private mSS2 As Reports.Spreadsheets

    Private msNew As TFRSpreadsheets.TFRCommon
    Private mobjReports As New Reports.ReportsCollection
    Private mobjSelectedReport As Reports.ReportsItem
    Private mflgLoading As Boolean = False
    Private mSQLDataReader As SqlDataReader
    Private PrevGroupsType As Reports.ReportsItem.ClientCodeSelect = -1
    Private pPNR As New PNRHistory.PNRs
    Private pE33_Gaslog As GaslogReports.E33_Collection
    Private pE35_Gaslog As GaslogReports.E35_Collection
    Private pGDSImport As New GDSImport.GDSImportItems

    Private Sub FrmTFReports_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            mflgLoading = True
            lblVersion.Text = VersionText
            fraCustomer.Visible = False
            fraReportOptions.Visible = False
            PopulateReportSelector()
            PrepareFrames()
            PrepareOptions()
            ValidateOptions()
        Catch ex As Exception
        Finally
            mflgLoading = False
        End Try

    End Sub

    Private Sub PrepareOptions()

        dtpDate1From.Value = mobjReports.Date1From
        dtpDate1To.Value = mobjReports.Date1To
        dtpDate2From.Value = mobjReports.Date2From
        dtpDate2To.Value = mobjReports.Date2To

        If mobjReports.SelectedCustomer <> "" Then
            cmbCustomers.SelectedIndex = cmbCustomers.FindString(mobjReports.SelectedCustomer)
        ElseIf mobjSelectedReport IsNot Nothing AndAlso mobjSelectedReport.InitialClientCode <> "" Then
            mobjReports.SelectedCustomer = mobjSelectedReport.InitialClientCode
            optCustomer.Checked = True
            cmbCustomers.SelectedIndex = cmbCustomers.FindString(mobjSelectedReport.InitialClientCode)
        End If

        cmbCustomerGroup.SelectedIndex = cmbCustomerGroup.FindString(mobjReports.CustomerGroup)
        ' BSP Month Report by airline
        If mobjReports.MonthDomestic Then
            opt02Domestic.Checked = True
        Else
            opt02International.Checked = True
        End If

        If lstBSPMonth.Items.Count > 0 Then
            lstBSPMonth.SelectedIndex = 0
            mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem("BSPDate")
        End If
        If mobjReports.BSPMonthDate IsNot Nothing Then
            mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem("BSPDate")
        End If
        ' BSP Fornight Report by ticket
        If lstBSPFortnight.Items.Count > 0 Then
            lstBSPFortnight.SelectedIndex = 0
        End If

        If lstBSPFortnight.Items.Count > 0 AndAlso lstBSPFortnight.SelectedItem("BSPDate") IsNot Nothing Then
            mobjReports.BSPFortDate = lstBSPFortnight.SelectedItem("BSPDate")
        End If
        chkCheckOption1.Checked = mobjReports.BooleanOption1

    End Sub

    Private Sub LoadYears()
        Try

            If lstYear.Items.Count = 0 Then
                Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)

                Dim dscmd As New SqlDataAdapter(sql.TransactionYears)
                Dim ds As New DataSet

                dscmd.Fill(ds)

                lstYear.Items.Clear()

                For i As Integer = CInt(ds.Tables(0).Rows(0).Item("MaxYear")) To CInt(ds.Tables(0).Rows(0).Item("MinYear")) Step -1
                    lstYear.Items.Add(i)
                Next
                lstYear.SelectedIndex = 0

                lstMonth.Items.Clear()

                For i As Integer = 1 To 12
                    lstMonth.Items.Add(i)
                Next
                lstMonth.SelectedIndex = Month(Today) - 1
                mobjReports.ReportYear = CInt(lstYear.SelectedItem)
                ValidateOptions()
            End If
        Catch ex As Exception
            Throw New Exception("LoadYears()" & vbCrLf & ex.Message)
        End Try

    End Sub
    Private Sub LoadBSPMonths()

        If lstBSPMonth.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)

            Dim dscmd As New SqlDataAdapter(sql.BSPMonths)
            Dim ds As New DataSet

            dscmd.Fill(ds)
            lstBSPMonth.DataSource = ds.Tables(0)
            lstBSPMonth.DisplayMember = "BSPDate"
            If lstBSPMonth.Items.Count > 0 Then
                lstBSPMonth.SelectedIndex = 0
                If lstBSPMonth.SelectedItem IsNot Nothing Then
                    mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem("BSPDate")
                End If
            End If
            ValidateOptions()

        End If

    End Sub
    Private Sub LoadVerificationYears()

        If lstBSPMonth.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForceSQL18, mobjReports)
            Dim dscmd As New SqlDataAdapter(sql.VerificationYears)
            Dim ds As New DataSet

            dscmd.Fill(ds)
            lstBSPMonth.DataSource = ds.Tables(0)
            lstBSPMonth.DisplayMember = "BSPDate"

            ValidateOptions()

        End If

    End Sub
    Private Sub LoadBSPFortnights()

        If lstBSPFortnight.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)

            Dim dscmd As New SqlDataAdapter(sql.BSPForthnights)
            Dim ds As New DataSet

            dscmd.Fill(ds)
            lstBSPFortnight.DataSource = ds.Tables(0)
            lstBSPFortnight.DisplayMember = "BSPDate"

            ValidateOptions()

        End If

    End Sub
    Private Sub LoadAgents()
        If lstGroupList.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)
            Dim dscmd As New SqlDataAdapter(sql.AgentGroups)
            Dim ds As New DataSet
            dscmd.Fill(ds)
            lstGroupList.DataSource = ds.Tables(0)
            lstGroupList.DisplayMember = "AgentName"
            ValidateOptions()
        End If
    End Sub
    Private Sub PopulateReportSelector()

        Dim pSelectedItem As Integer = 0

        With trvReportSelector
            .Nodes.Clear()
            For Each pReport As Reports.ReportsItem In mobjReports.Values
                If Not pReport.Hidden Then
                    If Not .Nodes.ContainsKey(pReport.GroupName) Then
                        .Nodes.Add(pReport.GroupName, pReport.GroupName)
                    End If
                    .Nodes(pReport.GroupName).Nodes.Add(pReport.Index.ToString, pReport.ReportName)
                End If

            Next
            .Sort()
        End With

    End Sub
    Private Sub PopulateCustomers()

        If cmbCustomers.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)

            Dim dscmd As New SqlDataAdapter(sql.ClientList)
            Dim ds As New DataSet

            dscmd.Fill(ds)
            cmbCustomers.DataSource = ds.Tables(0)
            cmbCustomers.DisplayMember = "DispName"

            ValidateOptions()
        End If
    End Sub
    Private Sub PopulateCustomerGroups(ByVal GroupsType As Reports.ReportsItem.ClientCodeSelect)

        If PrevGroupsType <> GroupsType Or cmbCustomerGroup.DataSource Is Nothing Then
            Dim sql As New TFReportsSQL.StoredProcedures(Reports.ReportsCollection.DBConnection.TravelForce, mobjReports)

            Dim dscmd As SqlDataAdapter
            If GroupsType = Reports.ReportsItem.ClientCodeSelect.ClientCodeAndSeachefsGroup Then
                dscmd = New SqlDataAdapter(sql.ClientGroupsSeaChefs)
            Else
                dscmd = New SqlDataAdapter(sql.ClientGroupsAll)
            End If
            Dim ds As New DataSet

            dscmd.Fill(ds)
            cmbCustomerGroup.DataSource = ds.Tables(0)
            cmbCustomerGroup.DisplayMember = "Description"
            ValidateOptions()
            PrevGroupsType = GroupsType
        End If
    End Sub

    Private Sub CmdExportExcel_Click(sender As Object, e As EventArgs) Handles cmdExportExcel.Click

        Me.Cursor = Cursors.WaitCursor

        Try
            Select Case mobjSelectedReport.Index
                Case 32, 33, 35, 39, 61, 62
                    SqlStatement(False)
                    ExportToExcel()
                Case Else
                    Dim dscmd As New SqlDataAdapter(SqlStatement(False))

                    ds = New DataSet
                    ds.Clear()
                    dscmd.Fill(ds)
                    ValidateOptions()
                    If ds.Tables(0).Rows.Count = 0 Then
                        MessageBox.Show("No transactions for the selected report and period")
                    Else
                        ExportToExcel()
                        ds.Dispose()
                    End If
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub
    Friend Function SqlStatement(ByVal withReader As Boolean) As SqlCommand

        If mobjSelectedReport IsNot Nothing Then
            SqlStatement = New SqlCommand
            Dim sql As New TFReportsSQL.StoredProcedures(mobjSelectedReport.DBConnection, mobjReports)
            With mobjReports
                Select Case mobjSelectedReport.Index
                    Case 0 ' Euronav
                        SqlStatement = sql.E00_Euronav()
                        lblTitle.Text = "00_Euronav " & .SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 1 ' Not Used
                    Case 2 ' BSP Month Report by airline
                        SqlStatement = sql.E02_BSPMonthReportByAirline()
                        lblTitle.Text = "02_BSP " & If(.MonthDomestic, "D ", "I ") & .BSPMonthDate
                    Case 3 ' BSP Fortnight Report by ticket
                        SqlStatement = sql.E03_BSPFortnightReportByTicket()
                        lblTitle.Text = "03_BSPTickets " & Format(.BSPFortDate, "yyyyMMdd")
                    Case 4 ' Ticket Info
                        SqlStatement = sql.E04_TicketInfo()
                        lblTitle.Text = "04_Ticket Info " & Format(Now, "yyyyMMdd-hhmm")
                    Case 5 ' Client turnover
                        SqlStatement = sql.E05_ClientTurnover()
                        lblTitle.Text = "05_Client Turnover - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 6 ' Profit per Client with BudgetComparison
                        SqlStatement = sql.E06_ProfitPerClientWithBudgetComparison()
                        lblTitle.Text = "06_Profit per Client with BudgetComparison " & Format(.ReportYear, "0000") & "-" & Format(.ReportMonth, "00")
                    Case 7 'Profit Per OPS Group
                        SqlStatement = sql.E07_ProfitPerOPSGroup()
                        lblTitle.Text = "07_Profit per OPS Group " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 8 'Profit Per Client Group
                        SqlStatement = sql.E08_ProfitPerClientGroup()
                        lblTitle.Text = "08_Profit per Client Group " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 9 'Profit Per OPS Group EXtra
                        SqlStatement = sql.E09_ProfitPerOPSGroupWithExtra()
                        lblTitle.Text = "09_Profit per OPS Group Extra " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 10 'Profit Per client Group EXtra
                        SqlStatement = sql.E10_ProfitPerClientGroupWithExtra()
                        lblTitle.Text = "10_Profit per Client Group Extra " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 11 'Profit Per Client Group with PY
                        SqlStatement = sql.E11_ProfitPerOPSGroupWithPY()
                        lblTitle.Text = "11_Profit per Client Group with PY " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 12 'Profit Per Client Group with Budget Comparison
                        'sqlStatement = pSQLCommands.E12_ProfitPerOPSGroupWithBudgetComparison()
                        SqlStatement = sql.E12_ProfitPerOPSGroupWithBudgetComparisonX(149)
                        lblTitle.Text = "12_Profit per Client Group with BudgetComparison " & Format(.ReportYear, "0000") & "-" & Format(.ReportMonth, "00")
                    Case 13 'Ticket Analysis
                        SqlStatement = sql.E13_TicketAnalysis()
                        lblTitle.Text = "13_Ticket Analysis " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 15 ' Daily Profit Report
                        SqlStatement = sql.E15_DailyProfitReportWithoutRINVA()
                        lblTitle.Text = "15_Daily Profit Report " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 16 ' Daily Profit Report with YTD
                        SqlStatement = sql.E16_DailyProfitReport()
                        lblTitle.Text = "16_Daily Profit Report WITH YTD " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 17 ' Service Fee Analysis
                        SqlStatement = sql.E17_ServiceFeeAnalysis()
                        lblTitle.Text = "17_Service fee Analysis " & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 18
                        SqlStatement = sql.E18_AirTicketSales()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "18_Air Ticket Sales All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "18_Air Ticket Sales " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "18_Air Ticket Sales " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 19 ' Profit Report Invoices with TicketNumber
                        SqlStatement = sql.E19_DailyProfitReportInvoicesWithTicketNumber(False)
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "19_Profit Report Invoices with Ticket Number and IW All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "19_Profit Report Invoices with Ticket Number and IW " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "19_Profit Report Invoices with Ticket Number and IW Group " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 20
                        SqlStatement = sql.E20_HellasConfidence()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "20_Air Ticket Sales All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "20_Air Ticket Sales " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "20_Air Ticket Sales " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 21
                        SqlStatement = sql.E21_ReportByVerifiedUser()
                        lblTitle.Text = "21_Optimize Report for " & .GroupList & " - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 22 ' 22-Euronav
                        SqlStatement = sql.E22_Euronav()
                        lblTitle.Text = "22_Euronav " & .SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 23, 36 ' 23-Sea Chefs
                        If mobjSelectedReport.Index = 23 Then
                            SqlStatement = sql.E23_SeaChefsX()
                            If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                                lblTitle.Text = "23_Sea Chefs " & .SelectedCustomer
                            Else
                                lblTitle.Text = "23_Sea Chefs " & .CustomerGroup
                            End If
                        Else
                            SqlStatement = sql.E36_SeaChefs_AllUnits()
                            lblTitle.Text = "36_Sea Chefs_AllUnits"
                        End If
                        If .Date1Checked Then
                            lblTitle.Text &= "- Inv Date " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                        If .Date2Checked Then
                            lblTitle.Text &= "- Departure Date " & Format(.Date2From, "yyyyMMdd") & "-" & Format(.Date2To, "yyyyMMdd")
                        End If
                    Case 24 ' 24 Profit per Agent Totals
                        SqlStatement = sql.E24_ProfitPerAgentTotals()
                        lblTitle.Text = "24_Profit per Agent Totals" & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 25 ' 25 Profit per Agent Transactions
                        SqlStatement = sql.E25_ProfitPerAgentTransactions()
                        lblTitle.Text = "25_Profit per Agent Transactions" & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 28 ' 28 Optimization Savings
                        SqlStatement = sql.E28_OptimizationSavings()
                        lblTitle.Text = "28_Optimization Savings" & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 29 ' 23-Sea Chefs Detailed
                        SqlStatement = sql.E29_SeaChefsDetailed()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "29_Sea Chefs Detailed " & .SelectedCustomer
                        Else
                            lblTitle.Text = "29_Sea Chefs Detailed" & .CustomerGroup
                        End If
                        If .Date1Checked Then
                            lblTitle.Text &= "- Inv Date " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                        If .Date2Checked Then
                            lblTitle.Text &= "- Departure Date " & Format(.Date2From, "yyyyMMdd") & "-" & Format(.Date2To, "yyyyMMdd")
                        End If
                    Case 30
                        SqlStatement = sql.E30_AirTicketsFullDetails()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "30_Air Ticket Full Details All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "30_Air Ticket Full Details " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "30_Air Ticket Full Details " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 31 ' 31-Sea Chefs Check
                        mobjReports.BooleanOption1 = True
                        SqlStatement = sql.E31_SeaChefsStatementCheck()
                        lblTitle.Text = "31_Sea Chefs Check"
                        If .Date1Checked Then
                            lblTitle.Text &= "- Inv Date " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                        If .Date2Checked Then
                            lblTitle.Text &= "- Departure Date " & Format(.Date2From, "yyyyMMdd") & "-" & Format(.Date2To, "yyyyMMdd")
                        End If
                    Case 32 ' 32-QC Selling Price
                        pPNR.ReadPNRs(mobjReports.Date1From, mobjReports.Date1To)
                        If .Date1From = .Date2From Then
                            lblTitle.Text = "32_QC Selling Price-" & Format(.Date1From, "yyyyMMdd")
                        Else
                            lblTitle.Text = "32_QC Selling Price-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 33 ' Gaslog 33-012212
                        pE33_Gaslog = New GaslogReports.E33_Collection(mobjReports.Date1From, mobjReports.Date1To, "012212")
                        lblTitle.Text = "33_Gaslog - 012212 Invoice Preparation - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 35 ' Gaslog 35-012217
                        pE35_Gaslog = New GaslogReports.E35_Collection(mobjReports.Date1From, mobjReports.Date1To, "012217")
                        lblTitle.Text = "35_Gaslog - 012217 Invoice Preparation - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 37 ' Profit Report Invoices with IW for all clients
                        ' this is the same as 19 with the added option to show IW10 for all cleints whereas 19 shows only for Sea Chefs
                        SqlStatement = sql.E19_DailyProfitReportInvoicesWithTicketNumber(True)
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "37_Profit Report Invoices with Ticket Number and IW All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "37_Profit Report Invoices with Ticket Number and IW " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "37_Profit Report Invoices with Ticket Number and IW Group " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 38
                        ' same as 18 with a different layout
                        SqlStatement = sql.E18_AirTicketSales()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "38_Air Ticket Sales All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "38_Air Ticket Sales " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "38_Air Ticket Sales " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 39
                        pGDSImport.Read()
                        lblTitle.Text = "39 GDS Import Errors-" & Format(Now, "yyyyMMddHHmm")
                    Case 40 'Profit Per Agent with Budget Comparison
                        SqlStatement = sql.E40_ProfitPerAgentWithBudgetComparison(144)
                        lblTitle.Text = "E40 Profit Per Agent with Budget Comparison " & Format(.ReportYear, "0000") & "-" & Format(.ReportMonth, "00")
                    Case 41 ' Daily Profit Report with RINVA Analysis
                        SqlStatement = sql.E41_DailyProfitReportWithRINVAAnalysis()
                        lblTitle.Text = "41_Daily Profit Report with RINVA Analysis " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 42
                        SqlStatement = sql.E42_AirTicketsWithFC()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "42_Air Ticket With FC All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "42_Air Ticket With FC " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "42_Air Ticket With FC " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 43 ' Daily Profit Report with Provisional Analysis
                        SqlStatement = sql.E43_DailyProfitReportWithProvisionalAnalysis()
                        lblTitle.Text = "43_Daily Profit Report with Provisional Analysis " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 45
                        SqlStatement = sql.E45_AirTicketSalesAll()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 46 ' Profit Report Invoices Totals
                        SqlStatement = sql.E19a_ProfitReportInvoicesTotals()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "19a_ProfitReportInvoicesTotals All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "19a_ProfitReportInvoicesTotals " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "19a_ProfitReportInvoicesTotals " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 47 ' Profit Report Totals Only
                        SqlStatement = sql.E47_DailyProfitReportTotalsOnly()
                        lblTitle.Text = "47_Daily Profit Report Totals Only " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "47_Daily Profit Report Totals Only " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "47_Daily Profit Report Totals Only " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "47_Daily Profit Report Totals Only " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 48 ' Profit Report Totals Per Invoice
                        SqlStatement = sql.E48_DailyProfitReportTotalsPerInvoice()
                        lblTitle.Text = "48 Daily Profit Report Totals per Invoice " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "48 Daily Profit Report Totals per Invoice " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "48 Daily Profit Report Totals per Invoice " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "48 Daily Profit Report Totals per Invoice " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 49 ' 49 Optimization Monthly Report
                        SqlStatement = sql.E49_Optimization_Monthly_Report()
                        lblTitle.Text = "49 Optimization Monthly Report " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 50 ' 50 Optimization Annual Report by Month
                        SqlStatement = sql.E50_Optimization_Annual_Report_by_Month()
                        lblTitle.Text = "50 Optimization Annual Report by Month " & .BSPMonthDate
                    Case 51 ' Daily Profit Report with Provisional Analysis
                        SqlStatement = sql.E51_Daily_Profit_Totals_per_Category()
                        lblTitle.Text = "51_Daily_Profit_Totals_per_Category " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 52 ' Gaslog Monthly Statement
                        SqlStatement = sql.E52_Gaslog_Monthly_Statement()
                        lblTitle.Text = "52_Gaslog_Monthly_Statement " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 53 ' Sea Chefs Invoice by Departure Date
                        SqlStatement = sql.E53_SeaChefs_InvoiceByDepartureDate()
                        lblTitle.Text = "53_SeaChefs_InvoiceByDepartureDate"
                    Case 54
                        SqlStatement = sql.E54_Client_Statement()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "54_Client_Statement All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "54_Client_Statement " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "54_Client_Statement " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 55
                        SqlStatement = sql.E55_Safety_Statement()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "55_Safety_Statement All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "55_Safety_Statement " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "55_Safety_Statement " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 56
                        SqlStatement = sql.E56_Clients()
                        lblTitle.Text = mobjSelectedReport.ReportName
                    Case 57
                        SqlStatement = sql.E57_TUI_030366()
                        lblTitle.Text = mobjSelectedReport.ReportName & " - " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 59
                        SqlStatement = sql.E59_AirTicketsByInvoiceDate()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "59_AirTicketsByInvoiceDate All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "59_AirTicketsByInvoiceDate " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "59_AirTicketsByInvoiceDate " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 60
                        SqlStatement = sql.E55_Safety_Statement()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = "60_Report_for_Lowest_Class All Clients-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = "60_Report_for_Lowest_Class " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = "60_Report_for_Lowest_Class " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 61 ' Gaslog 61-012629 Crew
                        pE33_Gaslog = New GaslogReports.E33_Collection(mobjReports.Date1From, mobjReports.Date1To, "012629")
                        lblTitle.Text = "61_Gaslog - 012629 Crew - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 62 ' Gaslog 62-012629 Corporate
                        pE35_Gaslog = New GaslogReports.E35_Collection(mobjReports.Date1From, mobjReports.Date1To, "012629")
                        lblTitle.Text = "62_Gaslog - 012629 Corporate - " & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 63
                        SqlStatement = sql.E63_AirTicketSalesTemenos
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If

                    Case 64
                        SqlStatement = sql.E64_LowestClasses()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 65
                        SqlStatement = sql.E65_OpsSales()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = mobjSelectedReport.ReportName & " " & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If
                    Case 66
                        SqlStatement = sql.E66_PurchasesPerAirline()
                        lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                    Case 67
                        SqlStatement = sql.E67_Columbia()
                        If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & mobjReports.SelectedCustomer & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        Else
                            lblTitle.Text = mobjSelectedReport.ReportName & "-" & mobjReports.CustomerGroup & "-" & Format(.Date1From, "yyyyMMdd") & "-" & Format(.Date1To, "yyyyMMdd")
                        End If

                    Case Else
                        Throw New Exception("No report selected")
                End Select
                If withReader Then
                    mSQLDataReader = sql.Reader(SqlStatement)
                End If
            End With
        Else
            SqlStatement = Nothing
        End If

    End Function
    Private Sub ExportToExcel()
        Try
            Dim pResponse As String = ""
            If mobjSelectedReport IsNot Nothing Then
                savExcel.FileName = lblTitle.Text
                savExcel.Filter = "Excel files (*.xlsx)|*.xlsx"
                savExcel.DefaultExt = "xlsx"
                If savExcel.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then
                    Throw New Exception("Export to file cancelled by user")
                End If
                pbarCounter.Minimum = 0
                pbarCounter.Maximum = 1
                pbarCounter.Value = 0
                MySpreadSheet = New SpreadSheet(ds, lblTitle.Text)
                mSS2 = New Reports.Spreadsheets()
                msNew = New TFRSpreadsheets.TFRCommon(mobjReports, ds, lblTitle.Text, savExcel.FileName)
                Me.Cursor = Cursors.WaitCursor
                With mobjReports
                    Select Case mobjSelectedReport.Index
                        Case 0 ' Euronav
                            msNew.E00_Euronav()
                            'mSS.E00_Euronav(savExcel.FileName)
                        Case 1 ' Not Used
                        Case 2 ' BSP Month Report by airline
                            msNew.E02_BSPMonthReportbyairline()
                            'mSS.E02_BSPMonthReportbyairline(.BSPMonthDate, savExcel.FileName)
                        Case 3 ' BSP Fortnight Report by ticket
                            MySpreadSheet.E03_BSPFortnightReportbyticket(.BSPFortDate, savExcel.FileName)
                        Case 4 ' Ticket Info
                            MySpreadSheet.E04_TicketInfo(mobjReports, savExcel.FileName)
                        Case 5 ' Client Turnover
                            MySpreadSheet.E05_ClientTurnover(savExcel.FileName)
                        Case 6 ' Profit per Client with Budget Comparison
                            MySpreadSheet.E06_ProfitPerClientwithBudgetComparison(mobjReports, savExcel.FileName)
                        Case 7 ' Profit Per OPS group
                            MySpreadSheet.E07_ProfitPerOPSgroup(.Date1From, .Date1To, savExcel.FileName)
                        Case 8 ' Profit Per client group
                            MySpreadSheet.E08_ProfitPerclientgroup(.Date1From, .Date1To, savExcel.FileName)
                        Case 9 ' Profit Per ops group extra
                            MySpreadSheet.E09_ProfitPerOpsGroupExtra(.Date1From, .Date1To, .Date2From, .Date2To, savExcel.FileName)
                        Case 10 ' Profit Per client group extra
                            MySpreadSheet.E10_ProfitPerClientGroupExtra(.Date1From, .Date1To, savExcel.FileName)
                        Case 11 ' Profit Per client group with PY
                            MySpreadSheet.E11_ProfitPerclientgroupwithPY(.Date1From, .Date1To, .FromYTD, .ToYTD, .FromPYTD, .ToPYTD, savExcel.FileName)
                        Case 12, 40 ' Profit Per client group with BudgetComparison
                            MySpreadSheet.E12_ProfitPerclientgroupwithBudgetComparison(mobjReports, savExcel.FileName) ' .DateFrom, .DateTo, .FromYTD, .ToYTD, .FromPYTD, .ToPYTD, savExcel.FileName)
                        Case 13 ' Ticket Analysis
                            MySpreadSheet.E13_TicketAnalysis(savExcel.FileName)
                        Case 15
                            MySpreadSheet.E15_DailyProfitReport(mobjReports, savExcel.FileName)
                        Case 16
                            MySpreadSheet.E16_DailyProfitReportWithYTD(mobjReports, savExcel.FileName)
                        Case 17
                            MySpreadSheet.E17_ServiceFeeAnalysis(mobjReports, savExcel.FileName)
                        Case 18
                            MySpreadSheet.E18_AirTicketSales(mobjReports, savExcel.FileName)
                        Case 19, 37
                            MySpreadSheet.E19_DailyProfitReportInvoicesWithTicketNumber(mobjReports, savExcel.FileName)
                        Case 20
                            MySpreadSheet.E20_HellasConfidence(mobjReports, savExcel.FileName)
                        Case 21
                            MySpreadSheet.E21_ReportByVerifiedUser(mobjReports, savExcel.FileName)
                        Case 22 ' 22-Euronav
                            MySpreadSheet.E22_Euronav(mobjReports, savExcel.FileName)
                        Case 23 ' 23-SeaChefs
                            MySpreadSheet.E23_SeaChefs(mobjReports, savExcel.FileName)
                        Case 24
                            MySpreadSheet.E24_ProfitPerAgentTotals(mobjReports, savExcel.FileName)
                        Case 25
                            MySpreadSheet.E25_ProfitPerAgentTransactions(mobjReports, savExcel.FileName)
                        Case 28
                            MySpreadSheet.E28_OptimizationSavings(mobjReports, savExcel.FileName)
                        Case 29 ' 29-SeaChefs Detailed
                            MySpreadSheet.E29_SeaChefsDetailed(mobjReports, savExcel.FileName)
                        Case 30, 59
                            MySpreadSheet.E30_AirTicketsFullDetails(mobjReports, savExcel.FileName)
                        Case 31
                            MySpreadSheet.E29_SeaChefsDetailed(mobjReports, savExcel.FileName)
                        Case 32
                            pResponse = mSS2.E32_QCSellingPrice(pPNR, savExcel.FileName)
                        Case 33, 61
                            pResponse = mSS2.E33_012212(pE33_Gaslog, savExcel.FileName)
                        Case 35, 62
                            pResponse = mSS2.E35_012217(pE35_Gaslog, savExcel.FileName)
                        Case 36
                            MySpreadSheet.E36_SeaChefs_AllUnits(mobjReports, savExcel.FileName)
                        Case 38
                            MySpreadSheet.E38_AirTicketSales(mobjReports, savExcel.FileName)
                        Case 39
                            pResponse = mSS2.E39_GDS_Import_Errors(mobjReports, pGDSImport, savExcel.FileName)
                        Case 41
                            MySpreadSheet.E41_DailyProfitReportWithRINVAAnalysis(mobjReports, savExcel.FileName)
                        Case 42
                            MySpreadSheet.E42_AirTicketsWithFC(mobjReports, savExcel.FileName)
                        Case 43
                            MySpreadSheet.E43_DailyProfitReportWithProvisionalAnalysis(mobjReports, savExcel.FileName)
                        Case 45
                            MySpreadSheet.E45_AirTicketSalesAll(mobjReports, savExcel.FileName)
                        Case 46
                            MySpreadSheet.E19aProfitReportInvoicesTotal(mobjReports, savExcel.FileName)
                        Case 47
                            MySpreadSheet.E47_DailyProfitReportTotalsOnly(mobjReports, savExcel.FileName)
                        Case 48
                            MySpreadSheet.E48_DailyProfitReportTotalsPerInvoice(mobjReports, savExcel.FileName)
                        Case 49
                            MySpreadSheet.E49_Optimization_Monthly_Report(mobjReports, savExcel.FileName)
                        Case 50
                            MySpreadSheet.E50_Optimization_Annual_Report_by_Month(mobjReports, savExcel.FileName)
                        Case 51
                            MySpreadSheet.E51_Daily_Profit_Totals_per_Category(mobjReports, savExcel.FileName)
                        Case 52
                            MySpreadSheet.E52_Gaslog_Monthly_Statement()
                        Case 53
                            MySpreadSheet.E53_SeaChefs_InvoicesByDepartureDate(mobjReports, savExcel.FileName)
                        Case 54
                            MySpreadSheet.E54_Client_Statement(mobjReports, savExcel.FileName)
                        Case 55
                            MySpreadSheet.E55_Safety_Statement(mobjReports, savExcel.FileName)
                        Case 56
                            MySpreadSheet.E56_ClientsPerGroup(mobjReports, savExcel.FileName)
                        Case 57
                            msNew.E57_TUI_030366()
                        Case 60
                            MySpreadSheet.E60_Report_For_Lowest_Class(mobjReports, savExcel.FileName)
                        Case 63
                            MySpreadSheet.E63_AirTicketSalesTemenos(mobjReports, savExcel.FileName)
                        Case 64
                            MySpreadSheet.E64_LowestClasses(mobjReports, savExcel.FileName)
                        Case 65
                            MySpreadSheet.E65_OpsSales(mobjReports, savExcel.FileName)
                        Case 66
                            MySpreadSheet.E66_PurchasesPerAirline(mobjReports, savExcel.FileName)
                        Case 67
                            MySpreadSheet.E67_Columbia(mobjReports, savExcel.FileName)
                        Case Else
                            Throw New Exception("No report selected")
                    End Select
                End With
            End If
            If pResponse <> "" Then
                MessageBox.Show(pResponse)
            End If
            Dim pPath As String = System.IO.Path.GetDirectoryName(savExcel.FileName)
            Process.Start("explorer.exe", "/select," & Chr(34) & savExcel.FileName & Chr(34)) ' pPath)
            pbarCounter.Value = 0
        Catch ex As Exception
            MessageBox.Show("cmdExportToExcel()" & vbCrLf & ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Function CheckNullZero(ByVal Value As Object) As Boolean

        Try
            If IsDBNull(Value) Then
                Return True
            ElseIf CDec(Value) = 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return True
        End Try

    End Function

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub DtpDate1From_ValueChanged(sender As Object, e As EventArgs) Handles dtpDate1From.ValueChanged

        If Not mflgLoading Then
            mobjReports.Date1From = dtpDate1From.Value
            mobjReports.Date1Checked = dtpDate1From.Checked
            ValidateOptions()
        End If

    End Sub

    Private Sub DtpDate1To_ValueChanged(sender As Object, e As EventArgs) Handles dtpDate1To.ValueChanged

        If Not mflgLoading Then
            mobjReports.Date1To = dtpDate1To.Value
            If mobjSelectedReport.Date2Init = Reports.ReportsItem.DateInitValue.From5DaysAfterEOM Then
                mobjReports.Date2From = DateAdd(DateInterval.Day, 5, DateSerial(mobjReports.Date1To.Year, mobjReports.Date1To.Month + 1, 0))
                mflgLoading = True
                dtpDate2From.Value = mobjReports.Date2From
                mflgLoading = False
            End If
            ValidateOptions()
        End If

    End Sub
    Private Sub DtpDate2From_ValueChanged(sender As Object, e As EventArgs) Handles dtpDate2From.ValueChanged

        If Not mflgLoading Then
            mobjReports.Date2From = dtpDate2From.Value
            mobjReports.Date2Checked = dtpDate2From.Checked
            ValidateOptions()
        End If
    End Sub

    Private Sub DtpDate2To_ValueChanged(sender As Object, e As EventArgs) Handles dtpDate2To.ValueChanged
        If Not mflgLoading Then
            mobjReports.Date2To = dtpDate2To.Value
            ValidateOptions()
        End If
    End Sub
    Private Sub ValidateOptions()
        If mobjSelectedReport IsNot Nothing Then
            Dim pflgLoading As Boolean = mflgLoading
            mflgLoading = True
            If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Then
                optAllCustomers.Checked = True
            ElseIf mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                optCustomer.Checked = True
            Else
                optCustomerGroup.Checked = True
            End If
            If mobjReports.OptionTriplet = 0 Then
                optTriplet0.Checked = True
            ElseIf mobjReports.OptionTriplet = 1 Then
                optTriplet1.Checked = True
            ElseIf mobjReports.OptionTriplet = 2 Then
                optTriplet2.Checked = True
            End If
            mflgLoading = pflgLoading
            Select Case mobjSelectedReport.Index
                Case 2 ' BSP Report
                    cmdExportExcel.Enabled = (mobjReports.BSPMonthDate <> "")
                Case 3 ' BSP Fortnight
                    cmdExportExcel.Enabled = (mobjReports.BSPFortDate <> Date.MinValue)
                Case 4 ' Ticket Info
                    cmdExportExcel.Enabled = (mobjReports.TextEntryItemsCount > 0)
                Case 6 ' Profit per client with budget Comparison
                    cmdExportExcel.Enabled = (lstYear.SelectedIndex >= 0 And lstMonth.SelectedIndex > 0)
                Case 12 ' Profit with budget comparison
                    cmdExportExcel.Enabled = (lstYear.SelectedIndex >= 0 And lstMonth.SelectedIndex >= 0)
                Case 17
                    cmdExportExcel.Enabled = (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients) Or ((mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient And mobjReports.SelectedCustomer <> "") Or (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByGroup And mobjReports.TagID <> 0))
                Case 18
                    cmdExportExcel.Enabled = mobjReports.OptionTriplet >= 0 And mobjReports.OptionTriplet <= 2 And (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients Or ((mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient And mobjReports.SelectedCustomer <> "") Or (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByGroup And mobjReports.TagID <> 0)))
                Case 19
                    cmdExportExcel.Enabled = (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.AllClients) Or ((mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient And mobjReports.SelectedCustomer <> "") Or (mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByGroup And mobjReports.TagID <> 0))
                Case 21
                    cmdExportExcel.Enabled = True 'mobjReports.GroupList <> ""
                Case 22
                    If mobjReports.ByClient = Reports.ReportsCollection.ClientReportType.ByClient Then
                        chkCheckOption1.Enabled = True
                    Else
                        chkCheckOption1.Enabled = False
                        chkCheckOption1.Checked = False
                    End If
                Case Else
                    cmdExportExcel.Enabled = (mobjReports.Date1Checked And mobjReports.Date1From <= mobjReports.Date1To) Or (mobjReports.Date2Checked And mobjReports.Date2From <= mobjReports.Date2To)
            End Select
        End If
    End Sub

    Private Sub CmbCustomers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCustomers.SelectedIndexChanged

        If Not mflgLoading Then
            mobjReports.SelectedCustomer = cmbCustomers.SelectedItem("Code")
            mflgLoading = True
            optCustomer.Checked = True
            mflgLoading = False
            ValidateOptions()
        End If

    End Sub
    Private Sub CmbCustomerGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCustomerGroup.SelectedIndexChanged

        If Not mflgLoading Then
            mobjReports.SetCustomerGroup(cmbCustomerGroup.SelectedItem("Id"), cmbCustomerGroup.SelectedItem("Description"))
            mflgLoading = True
            optCustomerGroup.Checked = True
            mflgLoading = False
            ValidateOptions()
        End If
    End Sub
    Private Sub PrepareFrames()

        Try

            Dim pLocX As Integer = 16
            Dim pLocY As Integer = 25
            Dim pStepX As Integer = 25
            Dim pStepY As Integer = 25
            '
            ' Options from Top to Bottom are:
            ' Customer Frame
            ' Options Frame
            ' in the Options frame we have
            ' Date1 From/To (with 3 labels for date description/from/to)
            ' Date2 From/To (with 3 labels for date description/from/to)
            ' CheckOption1
            ' opt02International
            ' opt02Domestic
            ' and then the list boxes/text boxes
            ' These are from left to right (and each one with its appropriate heading label):
            ' lstYear/lstMonth
            ' txtTicketNumbers
            ' lstBSPMonth
            ' lstBSPFortnight
            ' lstAirlines

            If mobjSelectedReport IsNot Nothing Then
                fraReportOptions.Visible = False
                With mobjSelectedReport
                    fraReportOptions.Text = .ReportName
                    If .ClientCode = Reports.ReportsItem.ClientCodeSelect.ClientCodeOnly Then
                        optAllCustomers.Visible = False
                        optCustomerGroup.Visible = False
                        cmbCustomerGroup.Visible = False
                        PopulateCustomers()
                        fraCustomer.Visible = True
                        optCustomer.Checked = True
                    ElseIf .ClientCode > 0 Then
                        optAllCustomers.Visible = True
                        optCustomerGroup.Visible = True
                        cmbCustomerGroup.Visible = True
                        PopulateCustomers()
                        PopulateCustomerGroups(.ClientCode)
                        fraCustomer.Visible = True
                    Else
                        fraCustomer.Visible = False
                    End If
                    If .Date1Text = "" Then
                        lblDate1.Visible = False
                        lblDate1To.Visible = False
                        dtpDate1From.Visible = False
                        dtpDate1To.Visible = False
                    Else
                        lblDate1.Text = .Date1Text
                        lblDate1.Visible = True
                        lblDate1.Location = New Point(pLocX, pLocY)
                        lblDate1To.Visible = True
                        lblDate1To.Location = New Point(lblDate1To.Location.X, pLocY)
                        dtpDate1From.ShowCheckBox = .Date1Optional
                        dtpDate1From.Checked = True
                        dtpDate1From.Visible = True
                        dtpDate1From.Location = New Point(dtpDate1From.Location.X, pLocY)
                        dtpDate1To.Visible = True
                        dtpDate1To.Location = New Point(dtpDate1To.Location.X, pLocY)
                        If .Date1Init = Reports.ReportsItem.DateInitValue.FromToPrevDayOrFriday Then
                            If Weekday(Today) = 2 Then
                                mobjReports.Date1From = DateAdd(DateInterval.Day, -3, Today)
                            Else
                                mobjReports.Date1From = DateAdd(DateInterval.Day, -1, Today)
                            End If
                            mobjReports.Date1To = DateAdd(DateInterval.Day, -1, Today)
                        ElseIf .Date1Init = Reports.ReportsItem.DateInitValue.FirstCurrMonthToToday Then
                            mobjReports.Date1Checked = True ' dtpDate1From.Checked
                            mobjReports.Date1From = DateSerial(Year(Today), Month(Today), 1) ' dtpDate1From.Value
                            mobjReports.Date1To = Today
                            If mobjReports.Date1To < mobjReports.Date1From Then
                                mobjReports.Date1To = mobjReports.Date1From
                            End If
                        ElseIf .Date1Init = Reports.ReportsItem.DateInitValue.FirstCurrMonthToYesterday Then
                            mobjReports.Date1Checked = True ' dtpDate1From.Checked
                            mobjReports.Date1From = DateSerial(Year(Today), Month(Today), 1) ' dtpDate1From.Value
                            mobjReports.Date1To = DateAdd(DateInterval.Day, -1, Today)
                            If mobjReports.Date1To < mobjReports.Date1From Then
                                mobjReports.Date1To = mobjReports.Date1From
                            End If
                        Else ' From/To Previous month 1st to last day of month
                            mobjReports.Date1From = (New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)).AddMonths(-1)
                            mobjReports.Date1To = mobjReports.Date1From.AddMonths(1).AddDays(-1)
                        End If
                        dtpDate1From.Value = mobjReports.Date1From
                        dtpDate1To.Value = mobjReports.Date1To
                        pLocY += pStepY
                    End If
                    If .Date2Text = "" Then
                        lblDate2.Visible = False
                        lblDate2To.Visible = False
                        dtpDate2From.Visible = False
                        dtpDate2To.Visible = False
                    Else
                        lblDate2.Text = .Date2Text
                        lblDate2.Visible = True
                        lblDate2.Location = New Point(pLocX, pLocY)
                        dtpDate2From.ShowCheckBox = .Date2Optional
                        dtpDate2From.Checked = True
                        dtpDate2From.Visible = True
                        dtpDate2From.Location = New Point(dtpDate2From.Location.X, pLocY)
                        If .Date2OnlyFrom Then
                            lblDate2To.Visible = False
                            dtpDate2To.Visible = False
                        Else
                            lblDate2To.Visible = True
                            lblDate2To.Location = New Point(lblDate2To.Location.X, pLocY)
                            dtpDate2To.Visible = True
                            dtpDate2To.Location = New Point(dtpDate2To.Location.X, pLocY)
                        End If
                        If .Date2Init = Reports.ReportsItem.DateInitValue.From5DaysAfterEOM Then
                            mobjReports.Date2From = DateAdd(DateInterval.Day, 5, DateSerial(mobjReports.Date1To.Year, mobjReports.Date1To.Month + 1, 0))
                            mobjReports.Date2To = mobjReports.Date2From
                        End If
                        dtpDate2From.Value = mobjReports.Date2From
                        dtpDate2To.Value = mobjReports.Date2To
                        pLocY += pStepY
                    End If
                    opt02Domestic.Visible = .DomInt
                    opt02International.Visible = .DomInt
                    If .DomInt Then
                        opt02Domestic.Location = New Point(pLocX, pLocY)
                        pLocY += pStepY
                        opt02International.Location = New Point(pLocX, pLocY)
                        pLocY += pStepY
                    End If
                    fraOptionsTriplet.Visible = .OptionsTriplet
                    If .OptionsTriplet Then
                        fraOptionsTriplet.Location = New Point(pLocX, pLocY)
                        optTriplet0.Text = .Options0Text
                        optTriplet1.Text = .Options1Text
                        optTriplet2.Text = .Options2Text
                        pLocY += pStepY + fraOptionsTriplet.Height
                    End If

                    If .CheckBoxText = "" Then
                        chkCheckOption1.Visible = False
                    Else
                        chkCheckOption1.Text = .CheckBoxText
                        chkCheckOption1.Visible = True
                        chkCheckOption1.Location = New Point(pLocX, pLocY)
                        pLocY += pStepY
                    End If
                    If .ReportYearMonth Then
                        LoadYears()
                        lblYear.Visible = True
                        lblYear.Location = New Point(pLocX, pLocY)
                        lstYear.Visible = True
                        lstYear.Location = New Point(pLocX, pLocY + lblYear.Size.Height)
                        lstYear.Height = fraReportOptions.Height - lstYear.Top - pStepY
                        lstYear.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstYear.Size.Width
                        lblMonth.Visible = True
                        lblMonth.Location = New Point(pLocX, pLocY)
                        lstMonth.Visible = True
                        lstMonth.Location = New Point(pLocX, pLocY + lblMonth.Size.Height)
                        lstMonth.Height = fraReportOptions.Height - lstMonth.Top - pStepY
                        lstMonth.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstMonth.Size.Width + pStepX
                    Else
                        lblYear.Visible = False
                        lblMonth.Visible = False
                        lstYear.Visible = False
                        lstMonth.Visible = False
                    End If
                    If .TextEntry <> "" Then
                        lblTextEntry.Text = .TextEntry
                        lblTextEntry.Visible = True
                        lblTextEntry.Location = New Point(pLocX, pLocY)
                        txtTextEntry.Visible = True
                        txtTextEntry.Multiline = .TextEntryMultiLine
                        txtTextEntry.Location = New Point(pLocX, pLocY + lblTextEntry.Size.Height)
                        txtTextEntry.Height = fraReportOptions.Height - txtTextEntry.Top - pStepY
                        txtTextEntry.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += txtTextEntry.Size.Width + pStepX
                    Else
                        lblTextEntry.Visible = False
                        txtTextEntry.Visible = False
                    End If
                    If .BSPMonth Then
                        LoadBSPMonths()
                        lblBSPMonth.Visible = True
                        lblBSPMonth.Location = New Point(pLocX, pLocY)
                        lblBSPMonth.Text = "BSP Month"
                        lstBSPMonth.Visible = True
                        lstBSPMonth.Location = New Point(pLocX, pLocY + lblBSPMonth.Size.Height)
                        lstBSPMonth.Height = fraReportOptions.Height - lstBSPMonth.Top - pStepY
                        lstBSPMonth.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstBSPMonth.Size.Width + pStepX
                    ElseIf .OptimizationMonth Then
                        LoadVerificationYears()
                        lblBSPMonth.Visible = True
                        lblBSPMonth.Location = New Point(pLocX, pLocY)
                        lblBSPMonth.Text = "Verified Year"
                        lstBSPMonth.Visible = True
                        lstBSPMonth.Location = New Point(pLocX, pLocY + lblBSPMonth.Size.Height)
                        lstBSPMonth.Height = fraReportOptions.Height - lstBSPMonth.Top - pStepY
                        lstBSPMonth.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstBSPMonth.Size.Width + pStepX
                    Else
                        lblBSPMonth.Visible = False
                        lstBSPMonth.Visible = False
                    End If
                    If .BSPFortnight Then
                        LoadBSPFortnights()
                        lblBSPFortnight.Visible = True
                        lblBSPFortnight.Location = New Point(pLocX, pLocY)
                        lstBSPFortnight.Visible = True
                        lstBSPFortnight.Location = New Point(pLocX, pLocY + lblBSPFortnight.Size.Height)
                        lstBSPFortnight.Height = fraReportOptions.Height - lstBSPFortnight.Top - pStepY
                        lstBSPFortnight.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstBSPFortnight.Size.Width + pStepX
                    Else
                        lblBSPFortnight.Visible = False
                        lstBSPFortnight.Visible = False
                    End If
                    If .GroupList <> Reports.ReportsCollection.GroupListType.Undefined Then
                        If .GroupList = Reports.ReportsCollection.GroupListType.OperatorsGroup Then
                            '                    loadOperationsGroup()
                        Else
                            LoadAgents()
                        End If
                        lblGroupList.Text = .GroupListText
                        lblGroupList.Visible = True
                        lblGroupList.Location = New Point(pLocX, pLocY)
                        lstGroupList.Visible = True
                        lstGroupList.Location = New Point(pLocX, pLocY + lblGroupList.Size.Height)
                        lstGroupList.Height = fraReportOptions.Height - lstGroupList.Top - pStepY
                        lstGroupList.Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom
                        pLocX += lstGroupList.Size.Width + pStepX
                    Else
                        lblGroupList.Visible = False
                        lstGroupList.Visible = False
                    End If
                End With
                fraReportOptions.Visible = True
            End If
        Catch ex As Exception
            Throw New Exception("PrepareFrames()" & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub OptInternational_CheckedChanged(sender As Object, e As EventArgs) Handles opt02International.CheckedChanged
        If Not mflgLoading Then
            mobjReports.MonthDomestic = False
            'prepareFrames()
            ValidateOptions()
        End If
    End Sub
    Private Sub OptDomestic_CheckedChanged(sender As Object, e As EventArgs) Handles opt02Domestic.CheckedChanged
        If Not mflgLoading Then
            mobjReports.MonthDomestic = True
            'prepareFrames()
            ValidateOptions()
        End If
    End Sub
    Private Sub LstBSPMonth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBSPMonth.SelectedIndexChanged
        If Not mflgLoading Then
            mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem("BSPDate")
            'prepareFrames()
            ValidateOptions()
        End If
    End Sub
    Private Sub LstBSPFortnight_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBSPFortnight.SelectedIndexChanged
        If Not mflgLoading Then
            With mobjReports
                .BSPFortDate = lstBSPFortnight.SelectedItem("BSPDate")
            End With
            'prepareFrames()
            ValidateOptions()
        End If
    End Sub
    Private Sub TrvReportSelector_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles trvReportSelector.AfterSelect
        Try

            If Not mflgLoading Then
                If e.Node.Parent Is Nothing Then
                    mobjSelectedReport = Nothing
                Else
                    mobjSelectedReport = mobjReports(CInt(e.Node.Name))
                End If
                Dim pFlagLoading As Boolean = mflgLoading
                mflgLoading = True
                mobjReports.ClearOptions()
                PrepareFrames()
                PrepareOptions()
                mflgLoading = pFlagLoading
                ValidateOptions()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub LstYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstYear.SelectedIndexChanged
        If Not mflgLoading Then
            With mobjReports
                .ReportYear = CInt(lstYear.SelectedItem)
            End With
            'prepareFrames()
            ValidateOptions()
        End If
    End Sub
    Private Sub LstMonth_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMonth.SelectedIndexChanged
        If Not mflgLoading Then
            With mobjReports
                .ReportMonth = CInt(lstMonth.SelectedItem)
            End With

            ValidateOptions()
        End If
    End Sub
    Private Sub ChkCheckOption1_CheckedChanged(sender As Object, e As EventArgs) Handles chkCheckOption1.CheckedChanged
        If Not mflgLoading Then
            With mobjReports
                .BooleanOption1 = chkCheckOption1.Checked
            End With

            ValidateOptions()
        End If
    End Sub
    Private Sub TxtTextEntry_TextChanged(sender As Object, e As EventArgs) Handles txtTextEntry.TextChanged
        With mobjReports
            .TextEntry = txtTextEntry.Text
        End With

        ValidateOptions()
    End Sub

    Private Sub OptTriplet0_CheckedChanged(sender As Object, e As EventArgs) Handles optTriplet0.CheckedChanged
        If Not mflgLoading Then
            mobjReports.OptionTriplet = 0
        End If
        ValidateOptions()
    End Sub

    Private Sub OptTriplet1_CheckedChanged(sender As Object, e As EventArgs) Handles optTriplet1.CheckedChanged
        If Not mflgLoading Then
            mobjReports.OptionTriplet = 1
        End If
        ValidateOptions()
    End Sub

    Private Sub OptTriplet2_CheckedChanged(sender As Object, e As EventArgs) Handles optTriplet2.CheckedChanged
        If Not mflgLoading Then
            mobjReports.OptionTriplet = 2
        End If
        ValidateOptions()
    End Sub

    Private Sub OptAllCustomers_CheckedChanged(sender As Object, e As EventArgs) Handles optAllCustomers.CheckedChanged
        If Not mflgLoading Then
            mobjReports.SetAllClients()
        End If
        ValidateOptions()
    End Sub

    Private Sub MySpreadSheet_ProgressCounter(CounterMin As Integer, CounterMax As Integer, CounterValue As Integer) Handles MySpreadSheet.ProgressCounter
        With pbarCounter
            .Minimum = CounterMin
            .Maximum = CounterMax
            .Value = CounterValue
        End With
    End Sub

    Private Sub LstGroupList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstGroupList.SelectedIndexChanged
        If Not mflgLoading Then
            mobjReports.GroupList = lstGroupList.Text
        End If
        ValidateOptions()
    End Sub

End Class
