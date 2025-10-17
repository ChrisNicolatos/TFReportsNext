Option Strict On
Option Explicit On
'Imports System.Data.SqlClient

' This is an old version so as not to lose the SQL commands

Public Class SQLCommandsOlder
    '    Private Const ConnectionStringACC As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=TravelForceCosmos;Persist Security Info=True;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    Private Const ConnectionStringPNR As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=AmadeusReports;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    'Private Const ConnectionStringACC As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=TravelForceCosmos;Persist Security Info=True;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    'Private Const ConnectionStringPNR As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=AmadeusReports;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    Dim mCnn As SqlConnection
    '    Public Sub New(ByVal DBConnection As ReportsCollection.DBConnection)
    '        MessageBox.Show(NextDB.DBConnections.TravelForcePanasoft)
    '        If DBConnection = ReportsCollection.DBConnection.TravelForce Then
    '            mCnn = New SqlConnection(ConnectionStringACC)
    '            'mCnn = New SqlConnection(ConnectionStringACC)
    '            mCnn = New SqlConnection(NextDB.DBConnections.TravelForcePanasoft)
    '            mCnn.Open()
    '        Else
    '            mCnn = New SqlConnection(ConnectionStringPNR)
    '            'mCnn = New SqlConnection(ConnectionStringPNR)
    '            mCnn = New SqlConnection(NextDB.DBConnections.TravelForcePanasoft)
    '            mCnn.Open()
    '        End If
    '    End Sub
    '    Public Function E00_Euronav(ByRef mReport As ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.CommandText = " SELECT ISNULL(Code + ' ' + Series,'') AS [Invoice Type] " &
    '                       "        , ISNULL([Number], '') AS [Invoice Number] " &
    '                       "        , ISNULL(CONVERT(VARCHAR,InvoiceDate,103), '') AS [Invoice Date] " &
    '                       "        , ISNULL(Department, '') AS [Vessel Name] " &
    '                       "        , ISNULL(DestinationAbbr ,'') AS [Destination] " &
    '                       "        , ISNULL(Currency, '') AS Currency " &
    '                       "        , ISNULL(SUM(Amount),0) AS Turnover " &
    '                       "        , ISNULL(BookedBy, '') AS BookedBy " &
    '                       "        , ISNULL(CPDepartment, '') AS CPDepartment " &
    '                       "        , ISNULL(MIN(Passengername), '') AS PassengerName " &
    '                       "        , ISNULL(MIN(Nationality),'') AS Nationality " &
    '                       "        , ISNULL(ClientCode, '') AS ClientCode " &
    '                       "        , ISNULL(PNRID, '') AS PNRID " &
    '                       "        , ISNULL(ConnectedDocument, '') AS ConnectedDocument " &
    '                       " FROM TravelForceCosmos.dbo.ViewCustomGriffinInvoicesGrouped " &
    '                       " WHERE (InvoiceDate BETWEEN @DateFrom AND @DateTo) AND ((ISNULL(@ClientCode, '') = '') OR ClientCode = @ClientCode) " &
    '                       " GROUP BY Code, Series, [Number], Department, DestinationAbbr, ClientCode, BookedBy, CPDepartment, InvoiceDate, Currency, PNRID, ConnectedDocument " &
    '                       " ORDER BY InvoiceDate, Code, Series, Number "

    '        Return sqlComm

    '    End Function
    '    Public Function E02_BSPMonthReportByAirline(ByVal Domestic As Boolean, ByVal BSPYear As Integer, ByVal BSPMonth As Integer) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand

    '        sqlComm.Parameters.Add("@Domestic", SqlDbType.Bit).Value = Domestic
    '        sqlComm.Parameters.Add("@BSPYear", SqlDbType.Int).Value = BSPYear
    '        sqlComm.Parameters.Add("@BSPMonth", SqlDbType.Int).Value = BSPMonth
    '        sqlComm.CommandText = " SELECT COALESCE( TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, 'Grand Total') AS IATAPrefix" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Airlines.IATACode, '') AS IATACode" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Airlines.AirlineName, '') AS AirlineName " &
    '                           " 	  , COALESCE( CONVERT(NCHAR(8), TravelForceCosmos.dbo.BSPTicket.Date, 112), '') AS BSPDate" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Currencies.ISOAlphabetic, '') AS Currency " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.CashTaxValue + TravelForceCosmos.dbo.BSPTicket.CreditTaxValue) AS [Tax]  " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.CashTransactionValue " &
    '                           "          + TravelForceCosmos.dbo.BSPTicket.CreditTransactionValue " &
    '                           " 		  + TravelForceCosmos.dbo.BSPTicket.CashTaxValue " &
    '                           "          + TravelForceCosmos.dbo.BSPTicket.CreditTaxValue) AS [FV]  " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.StandardCommissionAmount) AS Commission " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.SupplementaryDiscountAmount) AS Discount " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.Payable) AS Payable " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.VatOnCommission) AS VAT " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.Penalties) AS Penalties " &
    '                           " FROM TravelForceCosmos.dbo.BSPTicket LEFT OUTER JOIN " &
    '                           "     TravelForceCosmos.dbo.LookupTable ON TravelForceCosmos.dbo.BSPTicket.BSPTypeID = TravelForceCosmos.dbo.LookupTable.Id LEFT OUTER JOIN " &
    '                           "     TravelForceCosmos.dbo.Airlines ON TravelForceCosmos.dbo.BSPTicket.AirlineID = TravelForceCosmos.dbo.Airlines.Id " &
    '                           "     LEFT OUTER JOIN TravelForceCosmos.dbo.Currencies ON CurrencyID = Currencies.Id " &
    '                           " WHERE     (TravelForceCosmos.dbo.BSPTicket.Domestic = @Domestic) And YEAR(TravelForceCosmos.dbo.BSPTicket.Date) = @BSPYear And MONTH(TravelForceCosmos.dbo.BSPTicket.Date) = @BSPMonth " &
    '                           " GROUP BY TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic WITH ROLLUP " &
    '                           " HAVING GROUPING_ID(TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic) IN (0, 3 ,31)                       " &
    '                           " ORDER BY coalesce( TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, 'Grand Total'), IATACode, GROUPING_ID(TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic), ISOAlphabetic, BSPDate "
    '        Return sqlComm

    '    End Function
    '    Public Function E03_BSPFortnightReportByTicket(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@BSPDate", SqlDbType.Date).Value = mReport.BSPFortDate
    '        ' .CommandText = " SELECT ISNULL(Airlines.IATAAccountingPrefix, CONVERT(NVARCHAR(3),AirlineId)) + ' ' + ISNULL(Airlines.IATACode, '') + ' ' + ISNULL(Airlines.AirlineName, '') AS [Air] " 
    '        sqlComm.CommandText = " SELECT ISNULL(Airlines.IATAAccountingPrefix, '') + ' ' + ISNULL(Airlines.IATACode, '') + ' ' + ISNULL(Airlines.AirlineName, '') AS [Air] " &
    '                            " 	    ,CASE WHEN [Domestic] = 0 THEN 'I' ELSE 'D' END AS [I/D] " &
    '                            "       ,ISNULL(LookupTable.Name, 'UNKNOWN') AS Name " &
    '                            "       ,[DocumentNr] AS [Document Number] " &
    '                            "       ,[TransactionDate] AS [Issue Date] " &
    '                            "       ,ISNULL([CouponUseIndicator], 0) AS [CPUI] " &
    '                            "       ,ISNULL(Currencies.ISOAlphabetic, '') AS [Cur] " &
    '                            "       ,[CashTransactionValue] AS [Cash Transaction] " &
    '                            "       ,[CreditTransactionValue] AS [Credit Transaction] " &
    '                            "       ,[CashTaxValue] AS [Cash Tax] " &
    '                            "       ,[CreditTaxValue] AS [Credit Tax] " &
    '                            "       ,CONVERT(money,[StandardCommissionRate]) AS [Standard Commission Rate] " &
    '                            "       ,[StandardCommissionAmount] AS [Standard Commission Amount] " &
    '                            "       ,CONVERT(money,[SupplementaryDiscountRate]) AS [Supplementary Discount Rate] " &
    '                            "       ,[SupplementaryDiscountAmount] AS [Supplementary Discount Amount] " &
    '                            "       ,[VatOnCommission] AS [Tax on Commission] " &
    '                            "       ,[Payable] AS [Balance Payable] " &
    '                            "       ,ISNULL([Comments], '') AS [Comments] " &
    '                            "   FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                            "   LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines " &
    '                            "   ON AirlineID = Airlines.Id " &
    '                            "   LEFT OUTER JOIN [TravelForceCosmos].[dbo].[LookupTable] " &
    '                            "   ON BSPTypeID = LookupTable.Id " &
    '                            "   LEFT OUTER JOIN TravelForceCosmos.dbo.Currencies " &
    '                            "   ON CurrencyID = Currencies.Id " &
    '                            "   WHERE BSPTicket.Date = @BSPDate " &
    '                            "   ORDER BY  CASE WHEN [Domestic] = 0 THEN 'I' ELSE 'D' END, IATAAccountingPrefix,LookupTable.Code, DocumentNr "
    '        Return sqlComm
    '    End Function
    '    Public Function E04_TicketInfo(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand

    '        Dim pTickList As String = ""
    '        For i As Integer = 0 To mReport.TextEntryItemsCount
    '            If mReport.TextEntryItems(i).Length = 10 Then
    '                If pTickList.Length > 0 Then
    '                    pTickList &= ","
    '                End If
    '                pTickList &= "'" & mReport.TextEntryItems(i) & "'"
    '            End If
    '        Next

    '        If pTickList.Length > 0 Then

    '            sqlComm.CommandText = "SELECT    ISNULL(AirTickets.DocumentNr,'') AS DocumentNr " &
    '                                "   		, ISNULL(PNR.Code,'') AS Code " &
    '                                "   		, ISNULL(PNR.CreationPCC,'') AS CreationPCC " &
    '                                "   		, ISNULL(AirTickets.IssueDate,'01/01/01') AS IssueDate " &
    '                                "   		, ISNULL(GDSUsers.SignIn, '') AS SignIn " &
    '                                "   		, ISNULL(SalesPersons.Name, '') AS Salesperson " &
    '                                "   		, ISNULL(LLTicket.Name,'') AS TicketType " &
    '                                "   		, ISNULL(Airlines.IATACode,'') AS IATACode " &
    '                                "   		, ISNULL(AirTickets.Void,0) AS Void " &
    '                                "   		, ISNULL(ATEXch.DocumentNr, '') AS ExchangeTicket " &
    '                                "   		, ISNULL(TFEntities.Code,'') AS ClientCode " &
    '                                "   		, ISNULL(TFEntities.Name,'') AS ClientName " &
    '                                "   		, ISNULL(DocTypes.Code,'') + ISNULL(Documents.Series,'') + ISNULL(Documents.Number,'') AS [Invoice-CreditNote] " &
    '                                "   FROM AirTickets  " &
    '                                "   LEFT OUTER JOIN PNR  " &
    '                                "   	LEFT JOIN AirTickets ATExch " &
    '                                "   	ON ATExch.OriginalPNR = PNR.Code " &
    '                                "   ON AirTickets.PNRID = PNR.Id  " &
    '                                "   LEFT OUTER JOIN GDSUsers  " &
    '                                "   ON PNR.GDSUserID = GDSUsers.Id  " &
    '                                "   LEFT JOIN LookupTable LLTicket " &
    '                                "   ON AirTickets.AirTicketTypeID = LLTicket.Id " &
    '                                "   LEFT JOIN SalesPersons " &
    '                                "   ON GDSUsers.SalespersonID=SalesPersons.Id " &
    '                                "   LEFT JOIN Airlines " &
    '                                "   ON Airlines.Id = AirTickets.TicketingAirlineID " &
    '                                "   LEFT JOIN CommercialTransactions " &
    '                                "   	LEFT JOIN CommercialTransactionValues " &
    '                                "   		LEFT JOIN TFEntities " &
    '                                "   		ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID " &
    '                                "   		LEFT JOIN DocumentItems " &
    '                                "   			LEFT JOIN Documents " &
    '                                "   				LEFT JOIN DocTypes " &
    '                                "   				ON Documents.DocTypesID = DocTypes.Id " &
    '                                "   			ON Documents.Id = DocumentItems.DocumentsID " &
    '                                "   		ON DocumentItems.CommercialTransactionValueID=CommercialTransactionValues.Id " &
    '                                "   	ON CommercialTransactionValues.CommercialTransactionID=CommercialTransactions.Id AND CommercialTransactionValues.IsCost = 0 " &
    '                                "   ON CommercialTransactions.ProductNr = AirTickets.DocumentNr " &
    '                                "   WHERE AirTickets.DocumentNr IN (" & pTickList & ") " &
    '                                "   ORDER BY AirTickets.DocumentNr"
    '            Return sqlComm
    '        Else
    '            Throw New Exception("No tickets selected")
    '        End If

    '    End Function
    '    Public Function E05_ClientTurnover(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandText = " SELECT ClientCode AS ClientCode, ISNULL(TfEntities.Name, '') AS ClientName, ISNULL(Currency, '') AS Currency, SUM(ISNULL(Amount,0)) AS Turnover " &
    '                           " FROM TravelForceCosmos.dbo.ViewCustomGriffinInvoicesGrouped " &
    '                           " LEFT JOIN TravelForceCosmos.dbo.TfEntities " &
    '                           " ON ClientCode = TfEntities.Code " &
    '                           " WHERE (InvoiceDate BETWEEN @DateFrom AND @DateTo) " &
    '                           " GROUP BY ClientCode, TfEntities.Name, Currency "
    '        Return sqlComm

    '    End Function
    '    Public Function E06_ProfitPerClientWithBudgetComparison(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '        sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '        sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '        sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '        sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '        sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '        sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "  USE TravelForceCosmos   " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTablePYCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTablePYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableBudgetCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableBudgetYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWPYCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWPYtd  " &
    '                              "  End  " &
    '                              "    " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableClients  " &
    '                              "  End  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWCurr  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  	 ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWYTD  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWPYCurr  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWPYtd  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "    " &
    '                              "  SELECT  Code AS Client  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  " &
    '                              "        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  " &
    '                              "  	  ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  " &
    '                              "  INTO #TempTableBudgetCurr  " &
    '                              "  FROM [ATPIData].[dbo].[TFClientBudget]  " &
    '                              "  LEFT JOIN TravelForceCosmos.dbo.TFEntities  " &
    '                              "  ON tfbTFEntityId = TFEntities.Id  " &
    '                              "  WHERE tfbYear = @CurrYear AND tfbMonth = @ToMonth  " &
    '                              "  GROUP BY Code,Name  " &
    '                              "  ORDER BY Code  " &
    '                              "  SELECT  Code AS Client  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  " &
    '                              "        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  " &
    '                              "  	  ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  " &
    '                              "  INTO #TempTableBudgetYTD  " &
    '                              "  FROM [ATPIData].[dbo].[TFClientBudget]  " &
    '                              "  LEFT JOIN TravelForceCosmos.dbo.TFEntities  " &
    '                              "  ON tfbTFEntityId = TFEntities.Id  " &
    '                              "  WHERE tfbYear = @CurrYear AND tfbMonth BETWEEN @FromMonth AND @ToMonth  " &
    '                              "  GROUP BY Code,Name  " &
    '                              "  ORDER BY Code  " &
    '                              "     SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWCurr.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTableCurr  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWCurr  " &
    '                              "  	 ON #TempTableIWCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "   SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWYTD.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTableYTD  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWYTD  " &
    '                              "  	 ON #TempTableIWYTD.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "   	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "    SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYCurr.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTablePYCurr  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWPYCurr  " &
    '                              "  	 ON #TempTableIWPYCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "   SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYtd.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTablePYTD  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWPYtd  " &
    '                              "  	 ON #TempTableIWPYtd.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "  SELECT TFEntities.Code  " &
    '                              "  INTO #TempTableClients  " &
    '                              "  FROM TravelForceCosmos.dbo.TFEntities  " &
    '                              "  LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTablePYCurr     ON #TempTablePYCurr.Client     = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTablePYTD       ON #TempTablePYTD.Client       = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableBudgetCurr ON #TempTableBudgetCurr.Client = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableBudgetYTD  ON #TempTableBudgetYTD.Client  = TFEntities.Code  " &
    '                              "  WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL   " &
    '                              "      OR #TempTablePYTD.Client IS NOT NULL OR #TempTablePYCurr.Client IS NOT NULL   " &
    '                              "  	OR #TempTableBudgetCurr.Client IS NOT NULL OR #TempTableBudgetYTD.Client IS NOT NULL)  " &
    '                              "    AND (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0  " &
    '                              "      OR #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0  " &
    '                              "      OR #TempTablePYTD.Sales<>0 OR #TempTablePYTD.Profit <> 0 OR #TempTablePYTD.Pax <>0  " &
    '                              "      OR #TempTablePYCurr.Sales<>0 OR #TempTablePYCurr.Profit <> 0 OR #TempTablePYCurr.Pax <>0  " &
    '                              "      OR #TempTableBudgetCurr.Sales<>0 OR #TempTableBudgetCurr.Profit <> 0 OR #TempTableBudgetCurr.Pax <> 0  " &
    '                              "      OR #TempTableBudgetYTD.Sales<>0 OR #TempTableBudgetYTD.Profit <> 0 OR #TempTableBudgetYTD.Pax <> 0)  " &
    '                              "   SELECT  " &
    '                              "    'AllClients' AS GroupName  " &
    '                              "  	  , TFEntities.Code + '/' + TFEntities.Name AS Client  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  " &
    '                              "  	  , 0 AS ProfitPerPax  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Sales),0) AS SalesPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Profit),0) AS ProfitPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Pax),0) AS PaxPYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Sales),0) AS SalesPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Profit),0) AS ProfitPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Pax),0) AS PaxPYCurr  " &
    '                              "  	  , 0 AS ProfitPerPaxPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  " &
    '                              "  	  , 0 AS ProfitPerPaxBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxBudgetYTD  " &
    '                              "  FROM #TempTableClients  " &
    '                              "  LEFT JOIN TFEntities  " &
    '                              "  ON #TempTableClients.Code = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableCurr  " &
    '                              "  ON #TempTableCurr.Client=TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTablePYTD  " &
    '                              "  	ON #TempTablePYTD.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableYTD  " &
    '                              "  	ON #TempTableYTD.Client = TFEntities.Code	  " &
    '                              "   LEFT JOIN #TempTableBudgetCurr  " &
    '                              "  	ON #TempTableBudgetCurr.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableBudgetYTD  " &
    '                              "  	ON #TempTableBudgetYTD.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTablePYCurr  " &
    '                              "  	ON #TempTablePYCurr.Client = TFEntities.Code  " &
    '                              "   GROUP BY TFEntities.Code + '/' + TFEntities.Name  " &
    '                              "   ORDER BY TFEntities.Code + '/' + TFEntities.Name  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTablePYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTablePYCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableBudgetCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableBudgetYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableClients  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWPYCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWPYtd  " &
    '                              "   End  "

    '        Return sqlComm

    '    End Function
    '    Public Function E07_ProfitPerOPSGroup(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerOPSGroupStatement(False)

    '        Return sqlComm
    '    End Function
    '    Public Function E08_ProfitPerClientGroup(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerClientGroupStatement(False)

    '        Return sqlComm
    '    End Function
    '    Private Function ProfitPerClientGroupStatement(ByVal WithExtra As Boolean) As String

    '        Dim pReturn As String = ""
    '        pReturn &= " SELECT ISNULL(dbo.Tags.Description, '') AS GroupName  " &
    '" 		, ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name) AS Client" &
    '" 		,CONVERT(DECIMAL(18,2), SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END))                " &
    '" 					 AS Sales    " &
    '" 		,CONVERT(DECIMAL(18,2), SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END))                " &
    '" 					 AS Cost   " &
    '"         ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END)" &
    '" 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END)) AS Profit					  "

    '        If WithExtra Then
    '            pReturn &= " 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS FVExtra          " &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS TaxExtra" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Discount" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Commission" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS ServiceFee" &
    '" 		,CONVERT(DECIMAL(18,2),SUM(ISNULL(ServiceFeeAnalysis.Amount,0))) AS IW5" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate)) AS CancellationFee                "
    '        End If

    '        pReturn &= "		, SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax   " &
    '" 		, ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END)" &
    '" 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END))" &
    '" 		/(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax						" &
    '" FROM dbo.CommercialTransactions WITH (NOLOCK)    " &
    '" 	INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID   " &
    '" 	LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1   " &
    '" 	LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'    " &
    '" 	INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID    " &
    '" 	RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)     " &
    '" 		INNER JOIN dbo.Documents WITH (NOLOCK)      " &
    '" 			INNER JOIN dbo.TFEntities WITH (NOLOCK)      " &
    '" 				LEFT JOIN dbo.TFEntityTags        " &
    '" 					LEFT JOIN dbo.Tags ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID       " &
    '" 				ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)      " &
    '" 				LEFT JOIN dbo.TFEntityTags TFEntityTagsClientGroup        " &
    '" 					LEFT JOIN dbo.Tags TagClientGroup ON TagClientGroup.TagGroupID=146 AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID" &
    '" 				ON dbo.TFEntities.Id = TFEntityTagsClientGroup.TFEntityID AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=146 AND dbo.Tags.Id=TFEntityTagsClientGroup.TagID)      				" &
    '" 			ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id      " &
    '" 		ON dbo.DocTypes.Id = dbo.Documents.DocTypesID     " &
    '" 	ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id   " &
    '" WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'       " &
    '" 	  AND (dbo.Documents.IsCancellationDocument = 0)    " &
    '" 	  AND (dbo.Documents.DocStatusID = 41)   " &
    '" 	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)   " &
    '" 	  AND dbo.CommercialTransactionValues.Id IS NOT NULL     " &
    '" 	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '" GROUP BY ISNULL(dbo.Tags.Description, '')" &
    '" 		 , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)   " &
    '" ORDER BY ISNULL(dbo.Tags.Description, '')" &
    '" 		,ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)   "

    '        Return pReturn
    '    End Function
    '    Public Function E09_ProfitPerOPSGroupWithExtra(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerOPSGroupStatement(True)
    '        Return sqlComm
    '    End Function
    '    Private Function ProfitPerOPSGroupStatement(ByVal WithExtra As Boolean) As String

    '        ProfitPerOPSGroupStatement = ""
    '        ProfitPerOPSGroupStatement &= " SELECT ISNULL(dbo.Tags.Description, '') AS GroupName        " &
    '                            " 		, dbo.TFEntities.Code        " &
    '                            " 		, dbo.TFEntities.Name        " &
    '                            " 		, CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                            " 							ELSE 0 END))                " &
    '                            " 					 AS Sales    " &
    '                            " 		, CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                            " 							ELSE 0 END))                " &
    '                            " 					 AS Cost   " &
    '                            "         ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                            " 							ELSE 0 END)" &
    '                            " 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                            " 							ELSE 0 END)) AS Profit					  "
    '        If WithExtra Then
    '            ProfitPerOPSGroupStatement &= "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS FVExtra          " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS TaxExtra        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Discount        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Commission       " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS ServiceFee        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM(ISNULL(ServiceFeeAnalysis.Amount,0))) AS IW5                                                                                                                                             " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate)) AS CancellationFee  "

    '        End If
    '        ProfitPerOPSGroupStatement &= " 			, SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax   " &
    '                        " 		,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                        " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                        " 							ELSE 0 END)" &
    '                        " 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                        " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                        " 							ELSE 0 END))" &
    '                        " 		/(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax						" &
    '                        " FROM dbo.CommercialTransactions WITH (NOLOCK)    " &
    '                        " 	INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID   " &
    '                        " 	LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1   " &
    '                        " 	LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'    " &
    '                        " 	INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID    " &
    '                        " 	RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)     " &
    '                        " 		INNER JOIN dbo.Documents WITH (NOLOCK)      " &
    '                        " 			INNER JOIN dbo.TFEntities WITH (NOLOCK)      " &
    '                        " 				LEFT JOIN dbo.TFEntityTags        " &
    '                        " 					LEFT JOIN dbo.Tags         ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID       " &
    '                        " 				ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)      " &
    '                        " 			ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id      " &
    '                        " 		ON dbo.DocTypes.Id = dbo.Documents.DocTypesID     " &
    '                        " 	ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id   " &
    '                        " WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'       " &
    '                        " 	  AND (dbo.Documents.IsCancellationDocument = 0)    " &
    '                        " 	  AND (dbo.Documents.DocStatusID = 41)   " &
    '                        " 	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)   " &
    '                        " 	  AND dbo.CommercialTransactionValues.Id IS NOT NULL     " &
    '                        " 	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                        " GROUP BY ISNULL(dbo.Tags.Description, '')   " &
    '                        " 		 , dbo.TFEntities.Code  " &
    '                        " 		 , dbo.TFEntities.Name   " &
    '                        " ORDER BY ISNULL(dbo.Tags.Description, '')   " &
    '                        " 		,dbo.TFEntities.Code "
    '    End Function
    '    Public Function E10_ProfitPerClientGroupWithExtra(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerClientGroupStatement(True)
    '        Return sqlComm
    '    End Function
    '    Public Function E11_ProfitPerOPSGroupWithPY(ByVal TagGroup As Integer, ByVal DateFrom As Date, ByVal DateTo As Date, ByVal FromYTD As Date, ByVal ToYTD As Date, ByVal FromPY As Date, ByVal ToPY As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = DateFrom
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = DateTo
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = ToYTD
    '        sqlComm.Parameters.Add("@FromPY", SqlDbType.Date).Value = FromPY
    '        sqlComm.Parameters.Add("@ToPY", SqlDbType.Date).Value = ToPY
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = " USE TravelForceCosmos  " &
    '                            "   If(OBJECT_ID('tempdb..#TempTablePY') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTablePY   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableCurr   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableBudget') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableBudget   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableClients   " &
    '                            "   End  " &
    '                            "     SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTableCurr   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL      " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            "   SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTablePY   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromPY AND @ToPY)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL      " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            "   SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTableYTD   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL     " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            " SELECT TFEntities.Code " &
    '                            " INTO #TempTableClients " &
    '                            " FROM TFEntities " &
    '                            " LEFT JOIN #TempTableCurr ON #TempTableCurr.Client=TFEntities.Code " &
    '                            " LEFT JOIN #TempTableYTD ON #TempTableYTD.Client=TFEntities.Code " &
    '                            " LEFT JOIN #TempTablePY ON #TempTablePY.Client=TFEntities.Code " &
    '                            " WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL OR #TempTablePY.Client IS NOT NULL) " &
    '                            " 	  AND (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0 " &
    '                            " 	  OR   #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0 " &
    '                            " 	  OR   #TempTablePY.Sales<>0 OR #TempTablePY.Profit <> 0 OR #TempTablePY.Pax <>0) " &
    '                            "   SELECT  " &
    '                            "    ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED') AS GroupName        " &
    '                            "   	 , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name) AS Client      " &
    '                            "	     , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Profit)/NULLIF(SUM(#TempTableCurr.Pax),0),0) AS ProfitPerPax  " &
    '                            "	     , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Profit)/NULLIF(SUM(#TempTableYTD.Pax),0),0) AS ProfitPerPaxYTD   " &
    '                            "        , COALESCE(SUM(#TempTablePY.Sales),0) AS SalesPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Profit),0) AS ProfitPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Pax),0) AS PaxPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Profit)/NULLIF(SUM(#TempTablePY.Pax),0),0) AS ProfitPerPaxPY                                                 " &
    '                            "   FROM #TempTableClients         " &
    '                            " 	LEFT JOIN TFEntities " &
    '                            " 	ON #TempTableClients.Code = TFEntities.Code " &
    '                            "   		LEFT JOIN dbo.TFEntityTags                 " &
    '                            "   			LEFT JOIN dbo.Tags    " &
    '                            "   			ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID               " &
    '                            "   		ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID    " &
    '                            "   		   AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)              " &
    '                            "   	LEFT JOIN dbo.TFEntityTags TFEntityTagsClientGroup                 " &
    '                            "   		LEFT JOIN dbo.Tags TagClientGroup    " &
    '                            "   		ON TagClientGroup.TagGroupID=146    " &
    '                            "   			AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID        " &
    '                            "   	ON dbo.TFEntities.Id = TFEntityTagsClientGroup.TFEntityID    " &
    '                            "   		AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=146 AND dbo.Tags.Id=TFEntityTagsClientGroup.TagID)                 " &
    '                            "   LEFT JOIN #TempTableCurr " &
    '                            " 	ON #TempTableCurr.Client=TFEntities.Code   " &
    '                            "   LEFT JOIN #TempTablePY   " &
    '                            "   	ON #TempTablePY.Client = TFEntities.Code   " &
    '                            "   LEFT JOIN #TempTableYTD   " &
    '                            "   	ON #TempTableYTD.Client = TFEntities.Code	      " &
    '                            "   GROUP BY ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED') " &
    '                            "            , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)       " &
    '                            "   ORDER BY ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED')      " &
    '                            "            , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)     " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableCurr   " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTablePY') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTablePY   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD  " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableBudget') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD  " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableClients    " &
    '                            "   End "

    '        Return sqlComm
    '    End Function
    '    Public Function E12_ProfitPerOPSGroupWithBudgetComparison(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '        sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '        sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '        sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '        sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '        sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '        sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTablePYCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTablePYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableBudgetCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableBudgetYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWPYCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWPYtd  
    'End  

    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  
    'Begin  
    'Drop Table #TempTableClients  
    'End

    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWCurr  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '        ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '        ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    'GROUP BY CommercialTransactionValueID  

    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWYTD  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE  SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    'GROUP BY CommercialTransactionValueID  
    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWPYCurr  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )
    '        ) 		  



    'GROUP BY CommercialTransactionValueID  
    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWPYtd  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        ) 
    '        ) 
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        ) 
    '        ) 
    'GROUP BY CommercialTransactionValueID  

    'SELECT   ISNULL(Code, tfbOtherCode) AS Client  
    '        ,ISNULL(Name, tfbOtherName) AS ClientName
    '        ,tfbTFEntityId
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  
    '        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  
    '        ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  
    'INTO #TempTableBudgetCurr  
    'FROM [ATPIData].[dbo].[TFClientBudget]  
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities  
    'ON tfbTFEntityId = TFEntities.Id  
    'WHERE tfbYear = @CurrYear AND tfbMonth = @ToMonth  
    'GROUP BY ISNULL(Code, tfbOtherCode),ISNULL(Name, tfbOtherName),tfbTFEntityId  
    'ORDER BY ISNULL(Code, tfbOtherCode)  

    'SELECT   ISNULL(Code, tfbOtherCode) AS Client  
    '        ,ISNULL(Name, tfbOtherName) AS ClientName
    '        ,tfbTFEntityId
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  
    '        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  
    '        ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  
    'INTO #TempTableBudgetYTD  
    'FROM [ATPIData].[dbo].[TFClientBudget]  
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities  
    'ON tfbTFEntityId = TFEntities.Id  
    'WHERE tfbYear = @CurrYear AND tfbMonth BETWEEN @FromMonth AND @ToMonth  
    'GROUP BY ISNULL(Code, tfbOtherCode),ISNULL(Name, tfbOtherName),tfbTFEntityId  
    'ORDER BY ISNULL(Code, tfbOtherCode)  

    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWCurr.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTableCurr  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWCurr  
    '        ON #TempTableIWCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL 

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWYTD.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTableYTD  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWYTD  
    '        ON #TempTableIWYTD.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL  

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYCurr.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTablePYCurr  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWPYCurr  
    '        ON #TempTableIWPYCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYtd.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTablePYTD  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWPYtd  
    '        ON #TempTableIWPYtd.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    'SELECT TFEntities.Code  
    'INTO #TempTableClients  
    'FROM TravelForceCosmos.dbo.TFEntities  
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code  
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code  
    'LEFT JOIN #TempTablePYCurr     ON #TempTablePYCurr.Client     = TFEntities.Code  
    'LEFT JOIN #TempTablePYTD       ON #TempTablePYTD.Client       = TFEntities.Code  
    'LEFT JOIN #TempTableBudgetCurr ON #TempTableBudgetCurr.Client = TFEntities.Code  
    'LEFT JOIN #TempTableBudgetYTD  ON #TempTableBudgetYTD.Client  = TFEntities.Code  
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL   
    '    OR #TempTablePYTD.Client IS NOT NULL OR #TempTablePYCurr.Client IS NOT NULL   
    '    OR #TempTableBudgetCurr.Client IS NOT NULL OR #TempTableBudgetYTD.Client IS NOT NULL)  
    '    AND 
    '    (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0  
    '    OR #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0  
    '    OR #TempTablePYTD.Sales<>0 OR #TempTablePYTD.Profit <> 0 OR #TempTablePYTD.Pax <>0  
    '    OR #TempTablePYCurr.Sales<>0 OR #TempTablePYCurr.Profit <> 0 OR #TempTablePYCurr.Pax <>0  
    '    OR #TempTableBudgetCurr.Sales<>0 OR #TempTableBudgetCurr.Profit <> 0 OR #TempTableBudgetCurr.Pax <> 0  
    '    OR #TempTableBudgetYTD.Sales<>0 OR #TempTableBudgetYTD.Profit <> 0 OR #TempTableBudgetYTD.Pax <> 0)  

    '    SELECT  
    '    ISNULL(Tags.Description, '00 - UNCLASSIFIED') AS GroupName  
    '        , ISNULL(TagClientGroup.Description, TFEntities.Code + '/' + TFEntities.Name) AS Client  
    '        , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  
    '        , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  
    '        , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  
    '        , 0 AS ProfitPerPax  
    '        , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD  
    '        , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD  
    '        , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD  
    '        , 0 AS ProfitPerPaxYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Sales),0) AS SalesPYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Profit),0) AS ProfitPYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Pax),0) AS PaxPYTD  
    '        , 0 AS ProfitPerPaxPYTD  
    '        , COALESCE(SUM(#TempTablePYCurr.Sales),0) AS SalesPYCurr  
    '        , COALESCE(SUM(#TempTablePYCurr.Profit),0) AS ProfitPYCurr  
    '        , COALESCE(SUM(#TempTablePYCurr.Pax),0) AS PaxPYCurr  
    '        , 0 AS ProfitPerPaxPYCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  
    '        , 0 AS ProfitPerPaxBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  
    '        , 0 AS ProfitPerPaxBudgetYTD  
    'FROM #TempTableClients  
    'LEFT JOIN TFEntities  
    'ON #TempTableClients.Code = TFEntities.Code  
    '        LEFT JOIN TravelForceCosmos.dbo.TFEntityTags  
    '            LEFT JOIN TravelForceCosmos.dbo.Tags  
    '            ON Tags.TagGroupID=@TagGroup AND Tags.Id=dbo.TFEntityTags.TagID  
    '        ON TFEntities.Id = TFEntityTags.TFEntityID  
    '            AND TFEntityTags.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WHERE Tags.TagGroupID=@TagGroup AND Tags.Id=dbo.TFEntityTags.TagID)  
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup  
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup  
    '        ON TagClientGroup.TagGroupID=146  
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID  
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID  
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)  
    '    LEFT JOIN #TempTableCurr  
    '    ON #TempTableCurr.Client=TFEntities.Code  
    '    LEFT JOIN #TempTablePYTD  
    '    ON #TempTablePYTD.Client = TFEntities.Code  
    '    LEFT JOIN #TempTableYTD  
    '    ON #TempTableYTD.Client = TFEntities.Code	  
    '    LEFT JOIN #TempTableBudgetCurr  
    '    ON #TempTableBudgetCurr.Client = TFEntities.Code  
    '    LEFT JOIN #TempTableBudgetYTD  
    '    ON #TempTableBudgetYTD.Client = TFEntities.Code  
    '    LEFT JOIN #TempTablePYCurr  
    '    ON #TempTablePYCurr.Client = TFEntities.Code  
    '    GROUP BY ISNULL(Tags.Description, '00 - UNCLASSIFIED')  
    '            , ISNULL(TagClientGroup.Description, TFEntities.Code + '/' + TFEntities.Name)  

    'UNION

    'SELECT '99 - OTHER' AS GroupName  
    '     , #TempTableBudgetCurr.Client + '/' + #TempTableBudgetCurr.ClientName AS Client
    '        , 0 AS Sales  
    '        , 0 AS Profit  
    '        , 0 AS Pax  
    '        , 0 AS ProfitPerPax  
    '        , 0 AS SalesYTD  
    '        , 0 AS ProfitYTD  
    '        , 0 AS PaxYTD  
    '        , 0 AS ProfitPerPaxYTD  
    '        , 0 AS SalesPYTD  
    '        , 0 AS ProfitPYTD  
    '        , 0 AS PaxPYTD  
    '        , 0 AS ProfitPerPaxPYTD  
    '        , 0 AS SalesPYCurr  
    '        , 0 AS ProfitPYCurr  
    '        , 0 AS PaxPYCurr  
    '        , 0 AS ProfitPerPaxPYCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  
    '        , 0 AS ProfitPerPaxBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  
    '        , 0 AS ProfitPerPaxBudgetYTD  

    'FROM #TempTableBudgetCurr
    'LEFT JOIN #TempTableBudgetYTD
    'ON  #TempTableBudgetCurr.Client     = #TempTableBudgetYTD.Client
    'AND #TempTableBudgetCurr.ClientName = #TempTableBudgetYTD.ClientName
    'WHERE #TempTableBudgetCurr.tfbTFEntityId IS NULL
    'GROUP BY #TempTableBudgetCurr.Client + '/' + #TempTableBudgetCurr.ClientName
    '    ORDER BY GroupName, Client

    '    If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTablePYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTablePYCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableBudgetCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableBudgetYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableClients  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWPYCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWPYtd  
    '    End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E13_TicketAnalysis(ByVal UninvoicedOnly As Boolean, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromIssue", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToIssue", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromDep", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToDep", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "  USE TravelForceCosmos  
    '          SELECT dbo.TFEntities.Code AS ClientCode
    '        , dbo.TFEntities.Name AS ClientName
    '        , dbo.LookupTable.Name AS GDS
    '        , dbo.CommercialTransactions.[CreatorSalesmanString]
    '        , dbo.CommercialTransactions.[CreatorPCC]
    '        , dbo.CommercialTransactions.[IssueSalesmanString]
    '        , dbo.CommercialTransactions.[IssuePCC]
    '        , dbo.CommercialTransactions.[CustomDescription2] AS PNR
    '        , dbo.Airlines.IATAAccountingPrefix AS AirlineCode 
    '        , dbo.Airlines.IATACode AS AirlineAbbreviation
    '        , dbo.CommercialTransactions.[CustomDescription1] AS TicketNumber
    '        , dbo.AirTickets.IssueDate AS TicketIssueDate
    '        , ISNULL(dbo.AirTickets.NetRemitCode, '') AS TourCode
    '        , dbo.CommercialTransactions.[CustomDescription3] AS Passengers
    '        , dbo.CommercialTransactions.[CustomDescription4] AS Routing
    '        , dbo.CommercialTransactions.[FromDate] AS [Departure Date]
    '        , ISNULL(CPVBookedBy.Value, '') AS BookedBy      
    '        , ISNULL(CPVCostCentre.Value, '') AS CostCentre
    '        , ISNULL(dbo.TFEntityDepartments.Name, '') AS Vessel
    '        , ISNULL(CPVTrId.Value, '') AS TRID
    '        , ISNULL(dbo.Documents.IssueDate, '') AS InvoiceDate
    '        , ISNULL(dbo.DocTypes.Code, '') AS InvType
    '        , ISNULL(dbo.DocTypes.Series, '') AS InvSeries
    '        , ISNULL(dbo.Documents.Number, '') AS InvoiceNumber
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Fare          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Taxes 
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    
    '                                 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS CancFee 
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Discount          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       
    '                                 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Commission          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         
    '                                 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS ServiceFee          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       
    '                                 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         
    '                                 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    
    '                                 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Sales          
    '           , (ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax
    'FROM dbo.CommercialTransactions WITH (NOLOCK)         
    ' INNER JOIN dbo.CommercialTransactionValues    
    ' ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        
    '    AND dbo.CommercialTransactionValues.IsCost=0
    'INNER JOIN dbo.TFEntities WITH (NOLOCK)              
    'ON dbo.[CommercialTransactionValues].[CommercialEntityID] = dbo.TFEntities.Id
    'LEFT JOIN dbo.TFEntityDepartments
    'ON dbo.[CommercialTransactionValues].[CommercialEntityDepartmentID]= dbo.TFEntityDepartments.Id            
    'LEFT JOIN dbo.AirTickets
    '    ON dbo.CommercialTransactions.CustomDescription1 = dbo.AirTickets.DocumentNr
    'LEFT JOIN dbo.Airlines
    '    ON dbo.AirTickets.TicketingAirlineID = dbo.Airlines.Id		     
    'LEFT JOIN dbo.CustomPropertyValues CPVBookedBy
    '    ON dbo.AirTickets.BookFileId = CPVBookedBy.BookFileId 
    '        AND CPVBookedBy.CustomPropertyID=1 
    '        AND CPVBookedBy.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.CustomPropertyValues CPVCostCentre
    '    ON dbo.CommercialTransactionValues.CommercialTransactionID = CPVCostCentre.CTID 
    '        AND CPVCostCentre.CustomPropertyID=5 
    '        AND CPVCostCentre.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.CustomPropertyValues CPVTrId
    '    ON dbo.AirTickets.BookFileId = CPVTrId.BookFileId 
    '        AND CPVTrId.CustomPropertyID=14 
    '        AND CPVCostCentre.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.DocumentItems 
    '    ON dbo.DocumentItems.[CommercialTransactionValueID]= dbo.CommercialTransactionValues.Id  
    'LEFT JOIN dbo.Documents
    '    ON dbo.DocumentItems.DocumentsId = dbo.Documents.Id	    
    'LEFT JOIN dbo.DocTypes
    '    ON dbo.Documents.DocTypesId = dbo.DocTypes.Id
    'LEFT JOIN dbo.LookupTable
    '    ON dbo.AirTickets.GDSSystemId = dbo.LookupTable.Id
    ' WHERE    dbo.TFEntities.Code = @ClientCode  
    '          -- dbo.TFEntities.ID IN (SELECT TFEntityID FROM dbo.TFEntitytags WHERE TagId IN (154,155) )"
    '        If UninvoicedOnly Then
    '            sqlComm.CommandText &= " AND dbo.DocumentItems.Id IS NULL "
    '        End If
    '        sqlComm.CommandText &= "AND StatusID <>738 -- Not Cancelled
    '      AND dbo.CommercialTransactionValues.Id IS NOT NULL  
    '      AND dbo.AirTickets.IssueDate BETWEEN @FromIssue AND @ToIssue          
    'ORDER BY dbo.TFEntities.Code, dbo.TFEntities.Name, dbo.CommercialTransactions.[FromDate], dbo.AirTickets.IssueDate, dbo.CommercialTransactions.[CustomDescription2], dbo.CommercialTransactions.[CustomDescription1]   "
    '        Return sqlComm
    '    End Function
    '    Public Function E16_DailyProfitReport(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '        sqlComm.CommandTimeout = 400
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    'Begin
    'Drop Table #TempTableYTD
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWytd5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    'Begin
    'Drop Table #TempIWytd10
    'End
    'If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    'Begin
    'Drop Table #TempUninvoiced
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempUninvoiced -------------------------------------------------------------------------------------------------------------------
    ' SELECT TFEntities.Code AS Client
    '      , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate
    '            ELSE 0 END )) AS NetPayableUninvoiced
    '      , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS PaxUninvoiced
    ' INTO #TempUninvoiced
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID AND CommercialTransactionValues.IsCost = 0
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON CommercialTransactionValues.CommercialEntityID = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.AirTicketTransactions WITH (NOLOCK)
    'ON CommercialTransactions.Id = AirTicketTransactions.CommercialTransactionID
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND DocumentItems.Id IS NULL
    '      AND (CommercialTransactions.TransactionDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND StatusId = 331
    '      AND IsReversed = 0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    ' GROUP BY TFEntities.Code 
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0      
    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0      
    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempIWytd5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWytd10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableYTD --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableYTD
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWytd5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWytd10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code
    'LEFT JOIN #TempUninvoiced      ON #TempUninvoiced.Client      = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0
    '	OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0
    '    OR #TempUninvoiced.PaxUninvoiced >0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 1 AS Tots
    '      , ISNULL(TagClientGroup.Description, '') AS ClientGroupDescription
    '      , CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END AS ClientCode
    '      , ISNULL(TagClientGroup.Description, TFEntities.Name) AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' GROUP BY ISNULL(TagClientGroup.Description, ''),  CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END, ISNULL(TagClientGroup.Description, TFEntities.Name)

    'UNION

    '  SELECT 2 AS Tots
    '      , TagClientGroup.Description AS ClientGroupDescription
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' WHERE TagClientGroup.Description IS NOT NULL

    ' GROUP BY  TagClientGroup.Description, TFEntities.Code, TFEntities.Name

    ' ORDER BY Tots, Profit DESC
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    ' Begin
    ' Drop Table #TempTableYTD
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr10
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    ' Begin
    ' Drop Table #TempUninvoiced
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd10
    ' End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E15_DailyProfitReportWithoutRINVA(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '        sqlComm.CommandTimeout = 400
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    'Begin
    'Drop Table #TempTableYTD
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWytd5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    'Begin
    'Drop Table #TempIWytd10
    'End
    'If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    'Begin
    'Drop Table #TempUninvoiced
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempUninvoiced -------------------------------------------------------------------------------------------------------------------
    ' SELECT TFEntities.Code AS Client
    '      , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate
    '            ELSE 0 END )) AS NetPayableUninvoiced
    '      , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS PaxUninvoiced
    ' INTO #TempUninvoiced
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID AND CommercialTransactionValues.IsCost = 0
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON CommercialTransactionValues.CommercialEntityID = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.AirTicketTransactions WITH (NOLOCK)
    'ON CommercialTransactions.Id = AirTicketTransactions.CommercialTransactionID
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND DocumentItems.Id IS NULL
    '      AND (CommercialTransactions.TransactionDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND StatusId = 331
    '      AND IsReversed = 0
    ' GROUP BY TFEntities.Code 
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempIWytd5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWytd10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableYTD --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableYTD
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWytd5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWytd10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code
    'LEFT JOIN #TempUninvoiced      ON #TempUninvoiced.Client      = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0
    '	OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0
    '    OR #TempUninvoiced.PaxUninvoiced >0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 1 AS Tots
    '      , ISNULL(TagClientGroup.Description, '') AS ClientGroupDescription
    '      , CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END AS ClientCode
    '      , ISNULL(TagClientGroup.Description, TFEntities.Name) AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' GROUP BY ISNULL(TagClientGroup.Description, ''),  CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END, ISNULL(TagClientGroup.Description, TFEntities.Name)

    'UNION

    '  SELECT 2 AS Tots
    '      , TagClientGroup.Description AS ClientGroupDescription
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' WHERE TagClientGroup.Description IS NOT NULL

    ' GROUP BY  TagClientGroup.Description, TFEntities.Code, TFEntities.Name

    ' ORDER BY Tots, Profit DESC
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    ' Begin
    ' Drop Table #TempTableYTD
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr10
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    ' Begin
    ' Drop Table #TempUninvoiced
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd10
    ' End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E17_ServiceFeeAnalysis(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "USE TravelForceCosmos
    'SELECT dt.Code
    '     , dt.Series
    '     , d.Number
    '     , TFEntities.Code
    '     , TFEntities.Name
    '     , d.IssueDate
    '     , Currencies.ISOAlphabetic AS CurrencyCode
    '     , di.ServiceFee + di.ServiceFeeVatAmount AS InvoiceServiceFee
    '     , sfasum.Amount AS CardServFeeAmount
    '     , di.ServiceFee + di.ServiceFeeVatAmount - sfasum.Amount AS SFDifference
    '     , ISNULL(sfaTFee.Description, '') AS TfeeD
    '      , ISNULL(sfaTFee.Amount, 0) AS TFee
    '      , ISNULL(sfaTFeeDom.Description, '') AS TfeeDomD
    '      , ISNULL(sfaTFeeDom.Amount, 0) AS TFeeDom
    '      , ISNULL(sfaIW5.Description, '') AS IW5D
    '      , ISNULL(sfaIW5.Amount, 0) AS IW5
    '      , ISNULL(sfaIW6.Description, '') AS IW6D
    '      , ISNULL(sfaIW6.Amount, 0) AS IW6
    '      , ISNULL(sfaIW7.Description, '') AS IW7D
    '      , ISNULL(sfaIW7.Amount, 0) AS IW7
    '      , ISNULL(sfaIW8.Description, '') AS IW8D
    '      , ISNULL(sfaIW8.Amount, 0) AS IW8
    '      , ISNULL(sfaIW9.Description, '') AS IW9D
    '      , ISNULL(sfaIW9.Amount, 0) AS IW9
    '      , ISNULL(sfaIW10.Description, '') AS IW10D
    '      , ISNULL(sfaIW10.Amount, 0) AS IW10
    '      , ISNULL(sfaIW11.Description, '') AS IW11D
    '      , ISNULL(sfaIW11.Amount, 0) AS IW11
    'FROM Documents AS d 
    'INNER JOIN DocumentItems AS di 
    '    ON d.Id = di.DocumentsID 
    'INNER JOIN DocTypes AS dt 
    '    ON d.DocTypesID = dt.Id 
    'INNER JOIN (SELECT sfa.CommercialTransactionValueID, SUM(sfa.Amount) AS Amount
    '            FROM Documents AS d 
    '            INNER JOIN DocumentItems AS di 
    '                ON d.Id = di.DocumentsID 
    '            INNER JOIN ServiceFeeAnalysis AS sfa 
    '                ON di.CommercialTransactionValueID = sfa.CommercialTransactionValueID                               
    '            WHERE (d.IssueDate BETWEEN @FromDate AND @ToDate) 
    '              AND (d.DocStatusID = 41) 
    '              AND (d.IsCancellationDocument = 0)
    '            GROUP BY sfa.CommercialTransactionValueID) AS sfasum 
    '    ON di.CommercialTransactionValueID = sfasum.CommercialTransactionValueID 
    '    AND di.ServiceFee + di.ServiceFeeVatAmount <> sfasum.Amount
    'LEFT JOIN Currencies
    'ON d.CurrencyID =Currencies.Id
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW5
    'ON di.commercialtransactionvalueid = sfaIW5.commercialtransactionvalueid AND sfaIW5.ServiceFeeTypeId =1
    'LEFT JOIN ServiceFeeAnalysis AS sfaTFee
    'ON di.commercialtransactionvalueid = sfaTFee.commercialtransactionvalueid AND sfaTFee.ServiceFeeTypeId =2
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW6
    'ON di.commercialtransactionvalueid = sfaIW6.commercialtransactionvalueid AND sfaIW6.ServiceFeeTypeId =3
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW7
    'ON di.commercialtransactionvalueid = sfaIW7.commercialtransactionvalueid AND sfaIW7.ServiceFeeTypeId =4
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW8
    'ON di.commercialtransactionvalueid = sfaIW8.commercialtransactionvalueid AND sfaIW8.ServiceFeeTypeId =5
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW9
    'ON di.commercialtransactionvalueid = sfaIW9.commercialtransactionvalueid AND sfaIW9.ServiceFeeTypeId =6
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW10
    'ON di.commercialtransactionvalueid = sfaIW10.commercialtransactionvalueid AND sfaIW10.ServiceFeeTypeId =7
    'LEFT JOIN ServiceFeeAnalysis AS sfaTFeeDom
    'ON di.commercialtransactionvalueid = sfaTFeeDom.commercialtransactionvalueid AND sfaTFeeDom.ServiceFeeTypeId =8
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW11
    'ON di.commercialtransactionvalueid = sfaIW11.commercialtransactionvalueid AND sfaIW11.ServiceFeeTypeId =9
    'LEFT JOIN TFEntities
    'ON TFEntities.Id = d.CounterPartyID
    'WHERE (d.IsCancellationDocument = 0) 
    '  AND (d.DocStatusID = 41)
    'ORDER BY d.IssueDate

    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E18_AirTicketSales(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '	  , CASE WHEN ISNULL(Documents.IsCancellationDocument, 0) = 1 THEN ISNULL(DocTypes.CancelDocCode, '') ELSE ISNULL(DocTypes.Code, '') END AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    '      , ISNULL(Documents.DocStatusID, 0) AS DocStatusID
    '	  , ISNULL(CancDocType.Code + ' ' + CancDoc.Number, '') AS CancelsDoc
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    'LEFT JOIN TravelForceCosmos.dbo.Documents CancDoc WITH (NOLOCK)
    '	ON Documents.CancelsDocumentID = CancDoc.Id
    'LEFT JOIN TravelForceCosmos.dbo.DocTypes CancDocType WITH (NOLOCK)
    '    ON CancDocType.Id = CancDoc.DocTypesId
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND  CommercialTransactions.TransactionDate BETWEEN @FromDate AND @ToDate
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E19_DailyProfitReportInvoicesWithTicketNumber(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@WithTicket", SqlDbType.Bit).Value = mReport.BooleanOption1
    '        If mReport.ByClient = ReportsCollection.ClientReportType.AllClients Then
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '        ElseIf mReport.ByClient = ReportsCollection.ClientReportType.ByClient Then
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '        Else
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        End If

    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "USE TravelForceCosmos
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id 
    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id 

    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID)) 	
    '      AND TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) 
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '        , DocTypes.Code AS DocCode
    '        , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '        , Documents.Number as DocumentNumber
    '		, Documents.IssueDate
    '        , CASE WHEN @WithTicket = 1 THEN ISNULL(Airlines.IATACode, '') ELSE '' END AS Airline	 
    '        , CASE WHEN @WithTicket = 1 THEN CommercialTransactions.ProductNr ELSE '' END AS TicketNumber	 
    '        , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS FaceValueAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS TaxesAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS CommissionAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS DiscountAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS CancellationFeeAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS TFAIR


    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR


    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS FaceValueServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS TaxesServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS CommissionServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS DiscountServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS CancellationFeeServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS TFServices

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END) AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    '      , CommercialTransactionValues.Omit
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '     ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    'LEFT JOIN [TravelForceCosmos].[dbo].AirTickets ATTick WITH (NOLOCK)
    'ON ATTick.DocumentNr = CommercialTransactions.CustomDescription1
    'LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    'ON Airlines.Id = ATTick.TicketingAirlineID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID)) 	
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) 

    ' GROUP BY TFEntities.Code, TFEntityDepartments.Name, DocTypes.Code, Documents.Number, Documents.IssueDate, ISNULL(Airlines.IATACode, ''), CommercialTransactions.ProductNr, CommercialTransactionValues.Omit
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT DISTINCT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 
    '        TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , #TempTableCurr.Vessel
    '      , #TempTableCurr.IssueDate
    '      , #TempTableCurr.DocCode AS DocCode
    '      , #TempTableCurr.DocumentNumber AS DocumentNumber
    '      , #TempTableCurr.Airline AS Airline
    '      , ISNULL(#TempTableCurr.TicketNumber, '') AS TicketNumber
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueAir),0) AS FaceValueAir
    '      , COALESCE(SUM(#TempTableCurr.TaxesAir),0) AS TaxesAir
    '      , COALESCE(SUM(#TempTableCurr.CommissionAir),0) AS CommissionAir
    '      , COALESCE(SUM(#TempTableCurr.DiscountAir),0) AS DiscountAir
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeAir),0) AS CancellationFeeAir
    '      , COALESCE(SUM(#TempTableCurr.TFAir),0) AS TFAir
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxAir)),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueServices),0) AS FaceValueServices
    '      , COALESCE(SUM(#TempTableCurr.TaxesServices),0) AS TaxesServices
    '      , COALESCE(SUM(#TempTableCurr.CommissionServices),0) AS CommissionServices
    '      , COALESCE(SUM(#TempTableCurr.DiscountServices),0) AS DiscountServices
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeServices),0) AS CancellationFeeServices
    '      , COALESCE(SUM(#TempTableCurr.TFServices),0) AS TFServices

    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxServices)),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueAir),0) + COALESCE(SUM(#TempTableCurr.FaceValueServices),0) AS FaceValue
    '      , COALESCE(SUM(#TempTableCurr.TaxesAir),0) + COALESCE(SUM(#TempTableCurr.TaxesServices),0) AS Taxes
    '      , COALESCE(SUM(#TempTableCurr.CommissionAir),0) + COALESCE(SUM(#TempTableCurr.CommissionServices),0) AS Commission
    '      , COALESCE(SUM(#TempTableCurr.DiscountAir),0) + COALESCE(SUM(#TempTableCurr.DiscountServices),0) AS Discount
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeAir),0) + COALESCE(SUM(#TempTableCurr.CancellationFeeServices),0) AS CancellationFee
    '      , COALESCE(SUM(#TempTableCurr.TFAir),0) + COALESCE(SUM(#TempTableCurr.TFServices),0) AS TF

    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxAir)),0) + COALESCE(SUM(ABS(#TempTableCurr.PaxServices)),0),0),0)) AS ProfitPerPax
    '      , #TempTableCurr.Omit AS Omit
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' GROUP BY TFEntities.Code, TFEntities.Name, #TempTableCurr.Vessel, #TempTableCurr.DocCode, #TempTableCurr.DocumentNumber, #TempTableCurr.IssueDate, Airline, TicketNumber, #TempTableCurr.Omit
    ' ORDER BY #TempTableCurr.IssueDate, #TempTableCurr.DocCode, #TempTableCurr.DocumentNumber, TicketNumber
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableIWCurr
    ' End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E20_HellasConfidence(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromIssueDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToIssueDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@IssueDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromInvoiceDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToInvoiceDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@InvoiceDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , CommercialTransactions.CustomDescription4 AS Details
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(Documents.IssueDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate ) AS NetPayable      
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType	  
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber

    'FROM TravelForceCosmos.dbo.CommercialTransactionValues
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents
    '    ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.DocTypes
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    '    ON TFEntities.Id = CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  WHERE CommercialTransactionValues.IsCost = 0
    '        AND CommercialTransactionValues.Omit <> 1
    '        AND ISNULL(AirTickets.Void, 0) <> 1
    '        AND Documents.Id IS NOT NULL
    '        AND  (@IssueDateChecked = 0 OR CommercialTransactions.TransactionDate BETWEEN @FromIssueDate AND @ToIssueDate)
    '        AND  (@InvoiceDateChecked = 0 OR Documents.IssueDate BETWEEN @FromInvoiceDate AND @ToInvoiceDate)
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    'ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E21_ReportByVerifiedUser(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromVerifyDate", SqlDbType.DateTime).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToVerifyDate", SqlDbType.DateTime).Value = DateAdd(DateInterval.Hour, 24, mReport.Date1To)
    '        sqlComm.Parameters.Add("@VerifiedUserName", SqlDbType.NVarChar, 254).Value = mReport.GroupList
    '        sqlComm.CommandType = CommandType.Text
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#Temp') Is Not Null)
    'Begin
    'Drop Table #Temp
    'End

    'SELECT TFEntities.Code AS ClientCode, TFEntities.Name AS ClientName, Tags.[Description] AS OpsGroup
    'INTO #Temp
    'FROM      [TravelForceCosmos].[dbo].TFEntities AS TFEntities
    'LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityTags] AS TFEntityTags
    'LEFT JOIN [TravelForceCosmos].[dbo].Tags AS Tags
    'ON TagId = Tags.Id
    'ON TFEntities.Id = TFEntityTags.TFEntityId
    'WHERE TagGroupId = 149

    'USE AmadeusReports

    'SELECT ISNULL([doPCC], '') AS [PCC]
    '      ,ISNULL([doGDS], '') AS [GDS]
    '      ,ISNULL([doPNR], '') AS [PNR]
    '      ,ISNULL(PNRFinisherUsers.pfAgentName, doPCC + '-' + doUserGdsId) AS [Booked By]
    '      ,ISNULL([doDateLogged], '') AS [Date Logged]
    '      ,COALESCE(UsersVerified.pfAgentName, doVerifiedUserId, '') AS [Verified By]
    '      ,ISNULL([doVerifiedDate], '') AS [Date Verified]
    '      ,ISNULL([doVerificationReason], '') AS [Verification]
    '      ,ISNULL(DATEDIFF(MINUTE,doDateLogged, doVerifiedDate)/60.00, 0) AS HoursDiff
    '      ,ISNULL([doClientCode], '') AS [Client Code]
    '      ,ISNULL(#Temp.ClientName, '') AS [Client Name]
    '      ,ISNULL(#Temp.OpsGroup, '') AS [Client Group]
    '      ,[doPaxName] AS [Pax Name]
    '      ,[doItinerary] AS [Itinerary]
    '      ,[doTotal] AS [PNR Total]
    '      ,[doDownsellTotal] AS [Downsell Total]
    '      ,[doFareBasis] AS [PNR Fare Basis]
    '      ,[doDownsellFareBasis] AS [Downsell Fare Basis]
    '      ,[doGDSCommand] AS [GDS Pricing]
    '      ,[doIssuingCarrier] AS [Issuing Carrier]
    '  FROM [AmadeusReports].[dbo].[DownsellPNRLog]
    '  LEFT JOIN PNRFinisherUsers
    '  ON doPCC = pfPCC AND doUserGdsId = pfUser
    '  LEFT JOIN PNRFinisherUsers UsersVerified
    '  ON UsersVerified.pfPCC + '-' + UsersVerified.pfUser = doVerifiedUserId
    '  LEFT JOIN #Temp
    '  ON doClientCode = #Temp.ClientCode
    '  WHERE  ((doVerifiedUserId <> 'PRICER' AND  doVerificationReason <> 'NEW PRICER ENTRY')
    '            AND doVerifiedDate BETWEEN @FromVerifyDate AND @ToVerifyDate
    '            AND (@VerifiedUserName = '' OR ISNULL(UsersVerified.pfAgentName, doVerifiedUserId) = @VerifiedUserName))
    '  ORDER BY doVerifiedDate
    'If(OBJECT_ID('tempdb..#Temp') IS NOT NULL)
    'Begin
    'Drop Table #Temp
    'End
    '"
    '        Return sqlComm
    '    End Function
    '    Public Function E22_Euronav(ByRef mReport As ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND  Documents.IssueDate BETWEEN @FromDate AND @ToDate
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    '        AND (Documents.DocStatusID = 41)
    '        AND (Documents.DocTypesID NOT IN (74, 75))
    '	    AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '		AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E23_SeaChefs(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End
    'If(OBJECT_ID('tempdb..#TempBU') Is Not Null)
    'Begin
    'Drop Table #TempBU
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    '	  , TFEntities.Id AS EntityId
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes
    '    ON DocTypes.Id = Documents.DocTypesID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '	  AND (@TagID = 0 OR TFEntities.Code = @ClientCode OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    '      AND DocTypes.UsesFSD = 1
    'GROUP BY Documents.Id, TFEntities.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT DISTINCT #TempDate.EntityId, Tags.Description,
    'ISNULL(SeaChefsBusinessUnits.SupplierNumber, '') AS SupplierNumber,
    'ISNULL(SeaChefsBusinessUnits.SupplierSite, '') AS SupplierSite,
    'ISNULL(SeaChefsBusinessUnits.TaxRegNum, '') AS TaxRegNum
    'INTO #TempBU
    'FROM #TempDate
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityTags
    'ON TFEntityTags.TFEntityID=#TempDate.EntityId
    'LEFT JOIN TravelForceCosmos.dbo.Tags
    'ON TFEntityTags.TagID = Tags.Id
    'LEFT JOIN ATPIData.dbo.SeaChefsBusinessUnits
    'ON Tags.Description = SeaChefsBusinessUnits.BusinessUnit
    'WHERE Tags.TagGroupID=160

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL(#TempBU.Description, '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , ISNULL(#TempBU.SupplierNumber, '') AS [Supplier Number]
    '	  , ISNULL(#TempBU.SupplierSite, '') AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(#TempBU.TaxRegNum, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + CommercialTransactions.CustomDescription3 
    '	  + '_' + FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy')
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , CASE WHEN LEN(CPV05.Value)=26 THEN CPV05.Value + '-00000000-0000-0000-000000-000000' ELSE ISNULL(CPV05.Value, '') END AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	TFEntities.Code AS ClientCode
    '	  , TFEntities.Name AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , CommercialTransactions.CustomDescription2 AS PNR
    '	  , CommercialTransactions.CustomDescription1 AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , CommercialTransactionValues.ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ActionLT.Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , 'CY I STD EU SERVICE' AS TaxClassificationCode
    'FROM #TempDate WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.Documents
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN #TempBU WITH (NOLOCK)
    '    ON TFEntities.Id = #TempBU.EntityId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End
    'If(OBJECT_ID('tempdb..#TempBU') Is Not Null)
    'Begin
    'Drop Table #TempBU
    'End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E29_SeaChefsDetailed(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '	  AND (@TagID = 0 OR TFEntities.Code = @ClientCode OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    'GROUP BY Documents.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL((SELECT Tags.Description FROM TravelForceCosmos.dbo.TFEntityTags LEFT JOIN Tags ON TFEntityTags.TagID = Tags.ID
    '	    WHERE TFEntityTags.TFEntityID = TFEntities.ID and Tags.TagGroupID = 160 and Tags.ID BETWEEN 700 AND 706), '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , '300524' AS [Supplier Number]
    '	  , '300524-2' AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(TFEntities.TaxRegistrationCode, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription3, '') 
    '	  + '_' + ISNULL(FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy'), '') 
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , ISNULL(CPV05.Value, '') AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	ISNULL(TFEntities.Code, '') AS ClientCode
    '	  , ISNULL(TFEntities.Name, '') AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , ISNULL(CommercialTransactions.CustomDescription2, '') AS PNR
    '	  , ISNULL(Airlines.IATACode, '') AS AirlineCode
    '	  , ISNULL(CommercialTransactions.CustomDescription1, '') AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , ISNULL(CommercialTransactionValues.ProductTypeLT, '') AS ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(ActionLT.Id, 0) AS Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , ISNULL(CommercialTransactionValues.Remarks, '') AS Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    'FROM #TempDate WITH (NOLOCK)
    'LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End"
    '        Return sqlComm

    '    End Function
    '    Public Function E24_ProfitPerAgentTotals(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = 149
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT ISNULL(dbo.Tags.Description, '') AS GroupName   
    '     , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent  
    '	 , ISNULL(Salespersons.Name, '') as SalesPerson
    '     , dbo.TFEntities.Code                      
    '	 , dbo.TFEntities.Name                      
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Sales                  
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) 
    '	 + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Cost            
    '	 ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)             
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END)) AS Profit                                             
    '	 , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax                 
    '	 ,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                           
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))              
    '	 /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax                              
    '	 FROM dbo.CommercialTransactions WITH (NOLOCK)              
    '	 INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID  
    '	 LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '	 LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1             
    '	 LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'              
    '	 INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID              
    '	 RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)                   
    '	 INNER JOIN dbo.Documents WITH (NOLOCK)                        
    '	 INNER JOIN dbo.TFEntities WITH (NOLOCK)                            
    '	 LEFT JOIN dbo.TFEntityTags                                  
    '	 LEFT JOIN dbo.Tags         
    '	 ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID                             
    '	 ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)                        
    '	 ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id                    
    '	 ON dbo.DocTypes.Id = dbo.Documents.DocTypesID               
    '	 ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id    
    '	 WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'                   
    '	 AND (dbo.Documents.IsCancellationDocument = 0)                
    '	 AND (dbo.Documents.DocStatusID = 41)               
    '	 AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)               
    '	 AND dbo.CommercialTransactionValues.Id IS NOT NULL                 
    '	 AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL   
    '	 GROUP BY ISNULL(dbo.Tags.Description, '') 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 , dbo.TFEntities.Code                 
    '	 , dbo.TFEntities.Name    
    '	 ORDER BY ISNULL(dbo.Tags.Description, '')                 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 ,dbo.TFEntities.Code  "
    '        Return sqlComm

    '    End Function
    '    Public Function E25_ProfitPerAgentTransactions(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = 149
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT ISNULL(dbo.Tags.Description, '') AS GroupName   
    '     , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent  
    '	 , ISNULL(Salespersons.Name, '') as SalesPerson
    '     , dbo.TFEntities.Code                      
    '	 , dbo.TFEntities.Name    
    '	 , Documents.IssueDate
    ' 	 , ISNULL(DocTypes.Code, '') + ' ' + ISNULL(Documents.Series, '') + ' ' + ISNULL(Documents.Number, '') AS DocNumber
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Sales                  
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) 
    '	 + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Cost            
    '	 ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)             
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END)) AS Profit                                             
    '	 , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax                 
    '	 ,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                           
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))              
    '	 /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax                              
    '	 FROM dbo.CommercialTransactions WITH (NOLOCK)              
    '	 INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID  
    '	 LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '	 LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1             
    '	 LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'              
    '	 INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID              
    '	 RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)                   
    '	 INNER JOIN dbo.Documents WITH (NOLOCK)                        
    '	 INNER JOIN dbo.TFEntities WITH (NOLOCK)                            
    '	 LEFT JOIN dbo.TFEntityTags                                  
    '	 LEFT JOIN dbo.Tags         
    '	 ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID                             
    '	 ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)                        
    '	 ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id                    
    '	 ON dbo.DocTypes.Id = dbo.Documents.DocTypesID               
    '	 ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id    
    '	 WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'                   
    '	 AND (dbo.Documents.IsCancellationDocument = 0)                
    '	 AND (dbo.Documents.DocStatusID = 41)               
    '	 AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)               
    '	 AND dbo.CommercialTransactionValues.Id IS NOT NULL                 
    '	 AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL   
    '	 GROUP BY ISNULL(dbo.Tags.Description, '') 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 , dbo.TFEntities.Code                 
    '	 , dbo.TFEntities.Name
    '	 , Documents.IssueDate
    '	 , DocTypes.Code
    '	 , Documents.Series
    '	 , Documents.Number
    '	 ORDER BY ISNULL(dbo.Tags.Description, '')                 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 ,dbo.TFEntities.Code  "
    '        Return sqlComm

    '    End Function
    '    Public Function E28_OptimizationSavings(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT [doPNR] AS RecordLocator
    '      , '' AS TicketNumber
    '      ,[doPCC] AS PseudoCity
    '	  ,ISNULL(PNRFinisherUsers.pfAgentName, DownsellPNRLog.doUserGdsId) AS Consultant
    '      ,[doClientCode] AS AccountCode
    '	  ,ISNULL(TFEntities.Name, '') AS AccountName
    '	  ,SUBSTRING(doItinerary, 1,5) AS DateOfTravel
    '      ,SUBSTRING(doItinerary, 7,1000) AS IataItinerary
    '	  ,SUBSTRING(doItinerary, 11,2) AS PlateCarrier
    '	  , CASE WHEN CHARINDEX(' X ', doPaxName )>0 THEN
    '	    SUBSTRING(doPaxName, CHARINDEX(' X ', doPaxName) + 3, 1000)
    '		ELSE
    '	    LEN(REPLACE(doPaxName, '1.',''))-LEN(REPLACE(REPLACE(doPaxName, '1.',''), '.','')) + 1
    '		END AS NoOfPax
    '      ,[doTotal] AS FareAmount
    '	  ,[doVerifiedSavingsAmount] AS DownsellAmountPerPax
    '	  ,CASE WHEN CHARINDEX(' X ', doPaxName )>0 THEN
    '	    SUBSTRING(doPaxName, CHARINDEX(' X ', doPaxName) + 3, 1000)
    '		ELSE
    '	    LEN(REPLACE(doPaxName, '1.',''))-LEN(REPLACE(REPLACE(doPaxName, '1.',''), '.','')) + 1
    '		END * doVerifiedSavingsAmount AS DownsellAmountPerPNR
    '      , CASE WHEN doGDS = '1A' THEN DownsellPNRLog.doUserGdsId ELSE '' END AS AmadeusAgent
    '      , CASE WHEN doGDS = '1G' THEN DownsellPNRLog.doUserGdsId ELSE '' END AS GalileoAgent
    '      ,[doVerifiedUserId] AS ActionedBy
    '      ,[doFareBasis] AS OriginalFareBasis
    '      ,[doDownsellFareBasis] AS NewFareBasis
    '      ,[doGDSCommand] AS PricingCommand

    '  FROM [AmadeusReports].[dbo].[DownsellPNRLog]
    '  LEFT JOIN PNRFinisherUsers
    '  ON DownsellPNRLog.doPCC = PNRFinisherUsers.pfPCC AND DownsellPNRLog.doUserGdsId = PNRFinisherUsers.pfUser
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities
    '  ON doClientCode = TFEntities.Code
    '  WHERE doVerifiedsavingsamount >0
    '  AND CONVERT(DATE, doVerifiedDate) BETWEEN @FromDate AND @ToDate"
    '        Return sqlComm
    '    End Function
    '    Public Function E30_AirTicketsFullDetails(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@DateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked

    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPurchase
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS FaceValue
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Taxes
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Discount
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Commission
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS CancellationFee
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS FVExtra
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS TaxExtra
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS ServiceFee

    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND (@DateChecked=0 OR (CommercialTransactions.TransactionDate BETWEEN @FromDate AND @ToDate))
    '        AND (@DepDateChecked=0 OR (CommercialTransactions.FromDate BETWEEN @FromDepDate AND @ToDepDate))
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E31_SeaChefsStatementCheck(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN (154,155))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    'GROUP BY Documents.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL((SELECT Tags.Description FROM TravelForceCosmos.dbo.TFEntityTags LEFT JOIN Tags ON TFEntityTags.TagID = Tags.ID
    '	    WHERE TFEntityTags.TFEntityID = TFEntities.ID and Tags.TagGroupID = 160 and Tags.ID BETWEEN 700 AND 706), '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , '300524' AS [Supplier Number]
    '	  , '300524-2' AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(TFEntities.TaxRegistrationCode, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription3, '') 
    '	  + '_' + ISNULL(FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy'), '') 
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , ISNULL(CPV05.Value, '') AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	ISNULL(TFEntities.Code, '') AS ClientCode
    '	  , ISNULL(TFEntities.Name, '') AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , ISNULL(CommercialTransactions.CustomDescription2, '') AS PNR
    '	  , ISNULL(Airlines.IATACode, '')  AS AirlineCode
    '	  , ISNULL(CommercialTransactions.CustomDescription1, '') AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , ISNULL(CommercialTransactionValues.ProductTypeLT, '') AS ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(ActionLT.Id, 0) AS Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , ISNULL(CommercialTransactionValues.Remarks, '') AS Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '	  , ISNULL(Airlines.IATACode, '') AS AirlineCode
    'FROM #TempDate WITH (NOLOCK)
    'LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End"
    '        Return sqlComm

    '    End Function
    '    Public Function ClientList() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT Code, Name, [Code] + '-' + [Name] AS DispName  " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " WHERE IsClient=1  AND CanHaveCT=1 " &
    '                           " UNION Select Code, Name, [Name] + '-' + [Code] AS DispName  " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " WHERE IsClient=1  AND CanHaveCT=1 " &
    '                           " ORDER BY DispName"
    '        Return sqlComm

    '    End Function
    '    Public Function ClientGroupsAll() As SqlCommand

    '        ClientGroupsAll = New SqlCommand
    '        ClientGroupsAll = mCnn.CreateCommand
    '        ClientGroupsAll.CommandText = "SELECT [Id]
    '                            ,[Description]
    '                            FROM [TravelForceCosmos].[dbo].[Tags]
    '                            WHERE TagGroupID IN (146)
    '                            ORDER BY Description"
    '    End Function
    '    Public Function ClientGroupsSeaChefs() As SqlCommand

    '        ClientGroupsSeaChefs = New SqlCommand
    '        ClientGroupsSeaChefs = mCnn.CreateCommand
    '        ClientGroupsSeaChefs.CommandText = "SELECT [Id]
    '                            ,[Description]
    '                            FROM [TravelForceCosmos].[dbo].[Tags]
    '                            WHERE TagGroupID IN (162)
    '                            ORDER BY Description"
    '    End Function
    '    Public Function BSPMonths() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT LEFT(CONVERT(NCHAR(8), [Date], 112), 6) AS BSPDate " &
    '                           " FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                           " GROUP BY LEFT(CONVERT(NCHAR(8), [Date], 112), 6) " &
    '                           " ORDER BY BSPDate DESC"
    '        Return sqlComm

    '    End Function
    '    Public Function TransactionYears() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = "SELECT MIN(YEAR(TransactionDate)) AS MinYear, MAX(YEAR(TransactionDate)) AS MaxYear
    '                            FROM [TravelForceCosmos].[dbo].[CommercialTransactions]"
    '        Return sqlComm

    '    End Function

    '    Public Function BSPForthnights() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT Date AS BSPDate " &
    '                            " FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                            " GROUP BY [Date] " &
    '                            " ORDER BY BSPDate DESC"
    '        Return sqlComm

    '    End Function
    '    Public Function AgentGroups() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = "SELECT '' AS AgentName
    '                            UNION
    '                            SELECT DISTINCT pfAgentName AS AgentName
    '                            FROM [AmadeusReports].[dbo].[PNRFinisherUsers]
    '                            ORDER BY AgentName"
    '        Return sqlComm
    '    End Function
    '    Public Function Reader(cmm As SqlCommand) As SqlDataReader

    '        Reader = cmm.ExecuteReader

    '    End Function
    'End Class
    'Option Strict On
    'Option Explicit On
    'Imports System.Data.SqlClient
    'Public Class SQLCommands
    '    Private Const ConnectionStringACC As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=TravelForceCosmos;Persist Security Info=True;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    Private Const ConnectionStringPNR As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=AmadeusReports;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    'Private Const ConnectionStringACC As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=TravelForceCosmos;Persist Security Info=True;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    'Private Const ConnectionStringPNR As String = "Data Source=10.50.0.5\ATPICOSMOS,43334;Initial Catalog=AmadeusReports;User ID=tfsa;Password=aTPIprod!@#2020!"
    '    Dim mCnn As SqlConnection
    '    Public Sub New(ByVal DBConnection As ReportsCollection.DBConnection)
    '        MessageBox.Show(NextDB.DBConnections.TravelForcePanasoft)
    '        If DBConnection = ReportsCollection.DBConnection.TravelForce Then
    '            mCnn = New SqlConnection(ConnectionStringACC)
    '            'mCnn = New SqlConnection(ConnectionStringACC)
    '            mCnn = New SqlConnection(NextDB.DBConnections.TravelForcePanasoft)
    '            mCnn.Open()
    '        Else
    '            mCnn = New SqlConnection(ConnectionStringPNR)
    '            'mCnn = New SqlConnection(ConnectionStringPNR)
    '            mCnn = New SqlConnection(NextDB.DBConnections.TravelForcePanasoft)
    '            mCnn.Open()
    '        End If
    '    End Sub
    '    Public Function E00_Euronav(ByRef mReport As ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.CommandText = " SELECT ISNULL(Code + ' ' + Series,'') AS [Invoice Type] " &
    '                       "        , ISNULL([Number], '') AS [Invoice Number] " &
    '                       "        , ISNULL(CONVERT(VARCHAR,InvoiceDate,103), '') AS [Invoice Date] " &
    '                       "        , ISNULL(Department, '') AS [Vessel Name] " &
    '                       "        , ISNULL(DestinationAbbr ,'') AS [Destination] " &
    '                       "        , ISNULL(Currency, '') AS Currency " &
    '                       "        , ISNULL(SUM(Amount),0) AS Turnover " &
    '                       "        , ISNULL(BookedBy, '') AS BookedBy " &
    '                       "        , ISNULL(CPDepartment, '') AS CPDepartment " &
    '                       "        , ISNULL(MIN(Passengername), '') AS PassengerName " &
    '                       "        , ISNULL(MIN(Nationality),'') AS Nationality " &
    '                       "        , ISNULL(ClientCode, '') AS ClientCode " &
    '                       "        , ISNULL(PNRID, '') AS PNRID " &
    '                       "        , ISNULL(ConnectedDocument, '') AS ConnectedDocument " &
    '                       " FROM TravelForceCosmos.dbo.ViewCustomGriffinInvoicesGrouped " &
    '                       " WHERE (InvoiceDate BETWEEN @DateFrom AND @DateTo) AND ((ISNULL(@ClientCode, '') = '') OR ClientCode = @ClientCode) " &
    '                       " GROUP BY Code, Series, [Number], Department, DestinationAbbr, ClientCode, BookedBy, CPDepartment, InvoiceDate, Currency, PNRID, ConnectedDocument " &
    '                       " ORDER BY InvoiceDate, Code, Series, Number "

    '        Return sqlComm

    '    End Function
    '    Public Function E02_BSPMonthReportByAirline(ByVal Domestic As Boolean, ByVal BSPYear As Integer, ByVal BSPMonth As Integer) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand

    '        sqlComm.Parameters.Add("@Domestic", SqlDbType.Bit).Value = Domestic
    '        sqlComm.Parameters.Add("@BSPYear", SqlDbType.Int).Value = BSPYear
    '        sqlComm.Parameters.Add("@BSPMonth", SqlDbType.Int).Value = BSPMonth
    '        sqlComm.CommandText = " SELECT COALESCE( TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, 'Grand Total') AS IATAPrefix" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Airlines.IATACode, '') AS IATACode" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Airlines.AirlineName, '') AS AirlineName " &
    '                           " 	  , COALESCE( CONVERT(NCHAR(8), TravelForceCosmos.dbo.BSPTicket.Date, 112), '') AS BSPDate" &
    '                           "      , COALESCE( TravelForceCosmos.dbo.Currencies.ISOAlphabetic, '') AS Currency " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.CashTaxValue + TravelForceCosmos.dbo.BSPTicket.CreditTaxValue) AS [Tax]  " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.CashTransactionValue " &
    '                           "          + TravelForceCosmos.dbo.BSPTicket.CreditTransactionValue " &
    '                           " 		  + TravelForceCosmos.dbo.BSPTicket.CashTaxValue " &
    '                           "          + TravelForceCosmos.dbo.BSPTicket.CreditTaxValue) AS [FV]  " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.StandardCommissionAmount) AS Commission " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.SupplementaryDiscountAmount) AS Discount " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.Payable) AS Payable " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.VatOnCommission) AS VAT " &
    '                           "      , SUM(TravelForceCosmos.dbo.BSPTicket.Penalties) AS Penalties " &
    '                           " FROM TravelForceCosmos.dbo.BSPTicket LEFT OUTER JOIN " &
    '                           "     TravelForceCosmos.dbo.LookupTable ON TravelForceCosmos.dbo.BSPTicket.BSPTypeID = TravelForceCosmos.dbo.LookupTable.Id LEFT OUTER JOIN " &
    '                           "     TravelForceCosmos.dbo.Airlines ON TravelForceCosmos.dbo.BSPTicket.AirlineID = TravelForceCosmos.dbo.Airlines.Id " &
    '                           "     LEFT OUTER JOIN TravelForceCosmos.dbo.Currencies ON CurrencyID = Currencies.Id " &
    '                           " WHERE     (TravelForceCosmos.dbo.BSPTicket.Domestic = @Domestic) And YEAR(TravelForceCosmos.dbo.BSPTicket.Date) = @BSPYear And MONTH(TravelForceCosmos.dbo.BSPTicket.Date) = @BSPMonth " &
    '                           " GROUP BY TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic WITH ROLLUP " &
    '                           " HAVING GROUPING_ID(TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic) IN (0, 3 ,31)                       " &
    '                           " ORDER BY coalesce( TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, 'Grand Total'), IATACode, GROUPING_ID(TravelForceCosmos.dbo.Airlines.IATAAccountingPrefix, TravelForceCosmos.dbo.Airlines.IATACode, TravelForceCosmos.dbo.Airlines.AirlineName, TravelForceCosmos.dbo.BSPTicket.Date,  " &
    '                           "     TravelForceCosmos.dbo.Currencies.ISOAlphabetic), ISOAlphabetic, BSPDate "
    '        Return sqlComm

    '    End Function
    '    Public Function E03_BSPFortnightReportByTicket(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@BSPDate", SqlDbType.Date).Value = mReport.BSPFortDate
    '        ' .CommandText = " SELECT ISNULL(Airlines.IATAAccountingPrefix, CONVERT(NVARCHAR(3),AirlineId)) + ' ' + ISNULL(Airlines.IATACode, '') + ' ' + ISNULL(Airlines.AirlineName, '') AS [Air] " 
    '        sqlComm.CommandText = " SELECT ISNULL(Airlines.IATAAccountingPrefix, '') + ' ' + ISNULL(Airlines.IATACode, '') + ' ' + ISNULL(Airlines.AirlineName, '') AS [Air] " &
    '                            " 	    ,CASE WHEN [Domestic] = 0 THEN 'I' ELSE 'D' END AS [I/D] " &
    '                            "       ,ISNULL(LookupTable.Name, 'UNKNOWN') AS Name " &
    '                            "       ,[DocumentNr] AS [Document Number] " &
    '                            "       ,[TransactionDate] AS [Issue Date] " &
    '                            "       ,ISNULL([CouponUseIndicator], 0) AS [CPUI] " &
    '                            "       ,ISNULL(Currencies.ISOAlphabetic, '') AS [Cur] " &
    '                            "       ,[CashTransactionValue] AS [Cash Transaction] " &
    '                            "       ,[CreditTransactionValue] AS [Credit Transaction] " &
    '                            "       ,[CashTaxValue] AS [Cash Tax] " &
    '                            "       ,[CreditTaxValue] AS [Credit Tax] " &
    '                            "       ,CONVERT(money,[StandardCommissionRate]) AS [Standard Commission Rate] " &
    '                            "       ,[StandardCommissionAmount] AS [Standard Commission Amount] " &
    '                            "       ,CONVERT(money,[SupplementaryDiscountRate]) AS [Supplementary Discount Rate] " &
    '                            "       ,[SupplementaryDiscountAmount] AS [Supplementary Discount Amount] " &
    '                            "       ,[VatOnCommission] AS [Tax on Commission] " &
    '                            "       ,[Payable] AS [Balance Payable] " &
    '                            "       ,ISNULL([Comments], '') AS [Comments] " &
    '                            "   FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                            "   LEFT OUTER JOIN TravelForceCosmos.dbo.Airlines " &
    '                            "   ON AirlineID = Airlines.Id " &
    '                            "   LEFT OUTER JOIN [TravelForceCosmos].[dbo].[LookupTable] " &
    '                            "   ON BSPTypeID = LookupTable.Id " &
    '                            "   LEFT OUTER JOIN TravelForceCosmos.dbo.Currencies " &
    '                            "   ON CurrencyID = Currencies.Id " &
    '                            "   WHERE BSPTicket.Date = @BSPDate " &
    '                            "   ORDER BY  CASE WHEN [Domestic] = 0 THEN 'I' ELSE 'D' END, IATAAccountingPrefix,LookupTable.Code, DocumentNr "
    '        Return sqlComm
    '    End Function
    '    Public Function E04_TicketInfo(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand

    '        Dim pTickList As String = ""
    '        For i As Integer = 0 To mReport.TextEntryItemsCount
    '            If mReport.TextEntryItems(i).Length = 10 Then
    '                If pTickList.Length > 0 Then
    '                    pTickList &= ","
    '                End If
    '                pTickList &= "'" & mReport.TextEntryItems(i) & "'"
    '            End If
    '        Next

    '        If pTickList.Length > 0 Then

    '            sqlComm.CommandText = "SELECT    ISNULL(AirTickets.DocumentNr,'') AS DocumentNr " &
    '                                "   		, ISNULL(PNR.Code,'') AS Code " &
    '                                "   		, ISNULL(PNR.CreationPCC,'') AS CreationPCC " &
    '                                "   		, ISNULL(AirTickets.IssueDate,'01/01/01') AS IssueDate " &
    '                                "   		, ISNULL(GDSUsers.SignIn, '') AS SignIn " &
    '                                "   		, ISNULL(SalesPersons.Name, '') AS Salesperson " &
    '                                "   		, ISNULL(LLTicket.Name,'') AS TicketType " &
    '                                "   		, ISNULL(Airlines.IATACode,'') AS IATACode " &
    '                                "   		, ISNULL(AirTickets.Void,0) AS Void " &
    '                                "   		, ISNULL(ATEXch.DocumentNr, '') AS ExchangeTicket " &
    '                                "   		, ISNULL(TFEntities.Code,'') AS ClientCode " &
    '                                "   		, ISNULL(TFEntities.Name,'') AS ClientName " &
    '                                "   		, ISNULL(DocTypes.Code,'') + ISNULL(Documents.Series,'') + ISNULL(Documents.Number,'') AS [Invoice-CreditNote] " &
    '                                "   FROM AirTickets  " &
    '                                "   LEFT OUTER JOIN PNR  " &
    '                                "   	LEFT JOIN AirTickets ATExch " &
    '                                "   	ON ATExch.OriginalPNR = PNR.Code " &
    '                                "   ON AirTickets.PNRID = PNR.Id  " &
    '                                "   LEFT OUTER JOIN GDSUsers  " &
    '                                "   ON PNR.GDSUserID = GDSUsers.Id  " &
    '                                "   LEFT JOIN LookupTable LLTicket " &
    '                                "   ON AirTickets.AirTicketTypeID = LLTicket.Id " &
    '                                "   LEFT JOIN SalesPersons " &
    '                                "   ON GDSUsers.SalespersonID=SalesPersons.Id " &
    '                                "   LEFT JOIN Airlines " &
    '                                "   ON Airlines.Id = AirTickets.TicketingAirlineID " &
    '                                "   LEFT JOIN CommercialTransactions " &
    '                                "   	LEFT JOIN CommercialTransactionValues " &
    '                                "   		LEFT JOIN TFEntities " &
    '                                "   		ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID " &
    '                                "   		LEFT JOIN DocumentItems " &
    '                                "   			LEFT JOIN Documents " &
    '                                "   				LEFT JOIN DocTypes " &
    '                                "   				ON Documents.DocTypesID = DocTypes.Id " &
    '                                "   			ON Documents.Id = DocumentItems.DocumentsID " &
    '                                "   		ON DocumentItems.CommercialTransactionValueID=CommercialTransactionValues.Id " &
    '                                "   	ON CommercialTransactionValues.CommercialTransactionID=CommercialTransactions.Id AND CommercialTransactionValues.IsCost = 0 " &
    '                                "   ON CommercialTransactions.ProductNr = AirTickets.DocumentNr " &
    '                                "   WHERE AirTickets.DocumentNr IN (" & pTickList & ") " &
    '                                "   ORDER BY AirTickets.DocumentNr"
    '            Return sqlComm
    '        Else
    '            Throw New Exception("No tickets selected")
    '        End If

    '    End Function
    '    Public Function E05_ClientTurnover(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@Datefrom", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@DateTo", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandText = " SELECT ClientCode AS ClientCode, ISNULL(TfEntities.Name, '') AS ClientName, ISNULL(Currency, '') AS Currency, SUM(ISNULL(Amount,0)) AS Turnover " &
    '                           " FROM TravelForceCosmos.dbo.ViewCustomGriffinInvoicesGrouped " &
    '                           " LEFT JOIN TravelForceCosmos.dbo.TfEntities " &
    '                           " ON ClientCode = TfEntities.Code " &
    '                           " WHERE (InvoiceDate BETWEEN @DateFrom AND @DateTo) " &
    '                           " GROUP BY ClientCode, TfEntities.Name, Currency "
    '        Return sqlComm

    '    End Function
    '    Public Function E06_ProfitPerClientWithBudgetComparison(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '        sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '        sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '        sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '        sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '        sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '        sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "  USE TravelForceCosmos   " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTablePYCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTablePYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableBudgetCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableBudgetYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWYTD  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWPYCurr  " &
    '                              "  End  " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableIWPYtd  " &
    '                              "  End  " &
    '                              "    " &
    '                              "  If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  " &
    '                              "  Begin  " &
    '                              "  Drop Table #TempTableClients  " &
    '                              "  End  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWCurr  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  	 ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWYTD  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWPYCurr  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "  SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  " &
    '                              "  INTO #TempTableIWPYtd  " &
    '                              "  FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "  WHERE ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  " &
    '                              "    " &
    '                              "  SELECT DISTINCT CommercialTransactionValues.Id   " &
    '                              "    " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "   RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  " &
    '                              "   ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "  	  AND ServiceFeeAnalysis.Id IS NOT NULL  " &
    '                              "  	  AND CommercialTransactionValues.IsCost=0  " &
    '                              "  	  )  " &
    '                              "  GROUP BY CommercialTransactionValueID  " &
    '                              "    " &
    '                              "  SELECT  Code AS Client  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  " &
    '                              "        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  " &
    '                              "  	  ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  " &
    '                              "  INTO #TempTableBudgetCurr  " &
    '                              "  FROM [ATPIData].[dbo].[TFClientBudget]  " &
    '                              "  LEFT JOIN TravelForceCosmos.dbo.TFEntities  " &
    '                              "  ON tfbTFEntityId = TFEntities.Id  " &
    '                              "  WHERE tfbYear = @CurrYear AND tfbMonth = @ToMonth  " &
    '                              "  GROUP BY Code,Name  " &
    '                              "  ORDER BY Code  " &
    '                              "  SELECT  Code AS Client  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  " &
    '                              "        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  " &
    '                              "        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  " &
    '                              "  	  ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  " &
    '                              "  INTO #TempTableBudgetYTD  " &
    '                              "  FROM [ATPIData].[dbo].[TFClientBudget]  " &
    '                              "  LEFT JOIN TravelForceCosmos.dbo.TFEntities  " &
    '                              "  ON tfbTFEntityId = TFEntities.Id  " &
    '                              "  WHERE tfbYear = @CurrYear AND tfbMonth BETWEEN @FromMonth AND @ToMonth  " &
    '                              "  GROUP BY Code,Name  " &
    '                              "  ORDER BY Code  " &
    '                              "     SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWCurr.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTableCurr  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWCurr  " &
    '                              "  	 ON #TempTableIWCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "   SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWYTD.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTableYTD  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWYTD  " &
    '                              "  	 ON #TempTableIWYTD.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "   	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "    SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYCurr.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTablePYCurr  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWPYCurr  " &
    '                              "  	 ON #TempTableIWPYCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "   SELECT TFEntities.Code AS Client  " &
    '                              "  	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  " &
    '                              "  								 (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                              "  								 + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								 + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								 * CommercialTransactionValues.Rate  " &
    '                              "  								 ELSE 0 END))                       AS Sales  " &
    '                              "  	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  								  (ISNULL(CommercialTransactionValues.FaceValue, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.Taxes, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYtd.IWAmount,0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  " &
    '                              "  								  + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  " &
    '                              "  								  * CommercialTransactionValues.Rate  " &
    '                              "  								  ELSE 0 END)  " &
    '                              "  								  + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  " &
    '                              "  										(ISNULL(ctvCost.FaceValue, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FaceValueExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.FVXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.Taxes, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TaxesExtra, 0)  " &
    '                              "  										+ ISNULL(ctvCost.TAXXVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DiscountAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.DISCVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CommissionAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.COMVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.ServiceFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.SFVatAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CancellationFeeAmount, 0)  " &
    '                              "  										+ ISNULL(ctvCost.CFVatAmount, 0))  " &
    '                              "  										* ctvCost.Rate  " &
    '                              "  										ELSE 0 END)) AS Profit  " &
    '                              "  	  , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  " &
    '                              "   INTO #TempTablePYTD  " &
    '                              "   FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  " &
    '                              "  	 LEFT JOIN #TempTableIWPYtd  " &
    '                              "  	 ON #TempTableIWPYtd.CommercialTransactionValueID = CommercialTransactionValues.Id  " &
    '                              "   ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  " &
    '                              "   LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  " &
    '                              "   ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  " &
    '                              "  	AND ctvCost.IsCost=1  " &
    '                              "   INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  " &
    '                              "   ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  " &
    '                              "   RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  " &
    '                              "  	INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  " &
    '                              "  		INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  " &
    '                              "  		ON Documents.CounterPartyID = TFEntities.Id  " &
    '                              "  	ON DocTypes.Id = Documents.DocTypesID  " &
    '                              "   ON DocumentItems.DocumentsID = Documents.Id  " &
    '                              "   WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  " &
    '                              "  	  AND (Documents.IsCancellationDocument = 0)  " &
    '                              "  	  AND (Documents.DocStatusID = 41)  " &
    '                              "  	  AND (Documents.DocTypesID NOT IN (74, 75))  " &
    '                              "  	  AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  " &
    '                              "  	  AND CommercialTransactionValues.Id IS NOT NULL  " &
    '                              "  	  AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                              "   GROUP BY TFEntities.Code  " &
    '                              "   ORDER BY TFEntities.Code  " &
    '                              "  SELECT TFEntities.Code  " &
    '                              "  INTO #TempTableClients  " &
    '                              "  FROM TravelForceCosmos.dbo.TFEntities  " &
    '                              "  LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTablePYCurr     ON #TempTablePYCurr.Client     = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTablePYTD       ON #TempTablePYTD.Client       = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableBudgetCurr ON #TempTableBudgetCurr.Client = TFEntities.Code  " &
    '                              "  LEFT JOIN #TempTableBudgetYTD  ON #TempTableBudgetYTD.Client  = TFEntities.Code  " &
    '                              "  WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL   " &
    '                              "      OR #TempTablePYTD.Client IS NOT NULL OR #TempTablePYCurr.Client IS NOT NULL   " &
    '                              "  	OR #TempTableBudgetCurr.Client IS NOT NULL OR #TempTableBudgetYTD.Client IS NOT NULL)  " &
    '                              "    AND (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0  " &
    '                              "      OR #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0  " &
    '                              "      OR #TempTablePYTD.Sales<>0 OR #TempTablePYTD.Profit <> 0 OR #TempTablePYTD.Pax <>0  " &
    '                              "      OR #TempTablePYCurr.Sales<>0 OR #TempTablePYCurr.Profit <> 0 OR #TempTablePYCurr.Pax <>0  " &
    '                              "      OR #TempTableBudgetCurr.Sales<>0 OR #TempTableBudgetCurr.Profit <> 0 OR #TempTableBudgetCurr.Pax <> 0  " &
    '                              "      OR #TempTableBudgetYTD.Sales<>0 OR #TempTableBudgetYTD.Profit <> 0 OR #TempTableBudgetYTD.Pax <> 0)  " &
    '                              "   SELECT  " &
    '                              "    'AllClients' AS GroupName  " &
    '                              "  	  , TFEntities.Code + '/' + TFEntities.Name AS Client  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  " &
    '                              "  	  , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  " &
    '                              "  	  , 0 AS ProfitPerPax  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Sales),0) AS SalesPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Profit),0) AS ProfitPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYTD.Pax),0) AS PaxPYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxPYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Sales),0) AS SalesPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Profit),0) AS ProfitPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTablePYCurr.Pax),0) AS PaxPYCurr  " &
    '                              "  	  , 0 AS ProfitPerPaxPYCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  " &
    '                              "  	  , 0 AS ProfitPerPaxBudgetCurr  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  " &
    '                              "  	  , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  " &
    '                              "  	  , 0 AS ProfitPerPaxBudgetYTD  " &
    '                              "  FROM #TempTableClients  " &
    '                              "  LEFT JOIN TFEntities  " &
    '                              "  ON #TempTableClients.Code = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableCurr  " &
    '                              "  ON #TempTableCurr.Client=TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTablePYTD  " &
    '                              "  	ON #TempTablePYTD.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableYTD  " &
    '                              "  	ON #TempTableYTD.Client = TFEntities.Code	  " &
    '                              "   LEFT JOIN #TempTableBudgetCurr  " &
    '                              "  	ON #TempTableBudgetCurr.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTableBudgetYTD  " &
    '                              "  	ON #TempTableBudgetYTD.Client = TFEntities.Code  " &
    '                              "   LEFT JOIN #TempTablePYCurr  " &
    '                              "  	ON #TempTablePYCurr.Client = TFEntities.Code  " &
    '                              "   GROUP BY TFEntities.Code + '/' + TFEntities.Name  " &
    '                              "   ORDER BY TFEntities.Code + '/' + TFEntities.Name  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTablePYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTablePYCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableBudgetCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableBudgetYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableClients  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWYTD  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWPYCurr  " &
    '                              "   End  " &
    '                              "   If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  " &
    '                              "   Begin  " &
    '                              "   Drop Table #TempTableIWPYtd  " &
    '                              "   End  "

    '        Return sqlComm

    '    End Function
    '    Public Function E07_ProfitPerOPSGroup(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerOPSGroupStatement(False)

    '        Return sqlComm
    '    End Function
    '    Public Function E08_ProfitPerClientGroup(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerClientGroupStatement(False)

    '        Return sqlComm
    '    End Function
    '    Private Function ProfitPerClientGroupStatement(ByVal WithExtra As Boolean) As String

    '        Dim pReturn As String = ""
    '        pReturn &= " SELECT ISNULL(dbo.Tags.Description, '') AS GroupName  " &
    '" 		, ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name) AS Client" &
    '" 		,CONVERT(DECIMAL(18,2), SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END))                " &
    '" 					 AS Sales    " &
    '" 		,CONVERT(DECIMAL(18,2), SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END))                " &
    '" 					 AS Cost   " &
    '"         ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END)" &
    '" 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END)) AS Profit					  "

    '        If WithExtra Then
    '            pReturn &= " 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS FVExtra          " &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS TaxExtra" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Discount" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Commission" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS ServiceFee" &
    '" 		,CONVERT(DECIMAL(18,2),SUM(ISNULL(ServiceFeeAnalysis.Amount,0))) AS IW5" &
    '" 		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate)) AS CancellationFee                "
    '        End If

    '        pReturn &= "		, SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax   " &
    '" 		, ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '" 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '" 							ELSE 0 END)" &
    '" 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '" 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '" 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '" 							ELSE 0 END))" &
    '" 		/(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax						" &
    '" FROM dbo.CommercialTransactions WITH (NOLOCK)    " &
    '" 	INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID   " &
    '" 	LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1   " &
    '" 	LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'    " &
    '" 	INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID    " &
    '" 	RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)     " &
    '" 		INNER JOIN dbo.Documents WITH (NOLOCK)      " &
    '" 			INNER JOIN dbo.TFEntities WITH (NOLOCK)      " &
    '" 				LEFT JOIN dbo.TFEntityTags        " &
    '" 					LEFT JOIN dbo.Tags ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID       " &
    '" 				ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)      " &
    '" 				LEFT JOIN dbo.TFEntityTags TFEntityTagsClientGroup        " &
    '" 					LEFT JOIN dbo.Tags TagClientGroup ON TagClientGroup.TagGroupID=146 AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID" &
    '" 				ON dbo.TFEntities.Id = TFEntityTagsClientGroup.TFEntityID AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=146 AND dbo.Tags.Id=TFEntityTagsClientGroup.TagID)      				" &
    '" 			ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id      " &
    '" 		ON dbo.DocTypes.Id = dbo.Documents.DocTypesID     " &
    '" 	ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id   " &
    '" WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'       " &
    '" 	  AND (dbo.Documents.IsCancellationDocument = 0)    " &
    '" 	  AND (dbo.Documents.DocStatusID = 41)   " &
    '" 	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)   " &
    '" 	  AND dbo.CommercialTransactionValues.Id IS NOT NULL     " &
    '" 	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '" GROUP BY ISNULL(dbo.Tags.Description, '')" &
    '" 		 , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)   " &
    '" ORDER BY ISNULL(dbo.Tags.Description, '')" &
    '" 		,ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)   "

    '        Return pReturn
    '    End Function
    '    Public Function E09_ProfitPerOPSGroupWithExtra(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerOPSGroupStatement(True)
    '        Return sqlComm
    '    End Function
    '    Private Function ProfitPerOPSGroupStatement(ByVal WithExtra As Boolean) As String

    '        ProfitPerOPSGroupStatement = ""
    '        ProfitPerOPSGroupStatement &= " SELECT ISNULL(dbo.Tags.Description, '') AS GroupName        " &
    '                            " 		, dbo.TFEntities.Code        " &
    '                            " 		, dbo.TFEntities.Name        " &
    '                            " 		, CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                            " 							ELSE 0 END))                " &
    '                            " 					 AS Sales    " &
    '                            " 		, CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                            " 							ELSE 0 END))                " &
    '                            " 					 AS Cost   " &
    '                            "         ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                            " 							ELSE 0 END)" &
    '                            " 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                            " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                            " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                            " 							ELSE 0 END)) AS Profit					  "
    '        If WithExtra Then
    '            ProfitPerOPSGroupStatement &= "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS FVExtra          " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS TaxExtra        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Discount        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS Commission       " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))* dbo.CommercialTransactionValues.Rate)) AS ServiceFee        " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM(ISNULL(ServiceFeeAnalysis.Amount,0))) AS IW5                                                                                                                                             " &
    '                                        "		,CONVERT(DECIMAL(18,2),SUM((ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate)) AS CancellationFee  "

    '        End If
    '        ProfitPerOPSGroupStatement &= " 			, SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax   " &
    '                        " 		,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                        " 							(ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)          " &
    '                        " 							+ ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate               " &
    '                        " 							ELSE 0 END)" &
    '                        " 		+ SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN         " &
    '                        " 							(ISNULL(ctvCost.FaceValue, 0)              + ISNULL(ctvCost.FVVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.FaceValueExtra, 0)        + ISNULL(ctvCost.FVXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.Taxes, 0)                 + ISNULL(ctvCost.TAXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.TaxesExtra, 0)            + ISNULL(ctvCost.TAXXVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.DiscountAmount, 0)        + ISNULL(ctvCost.DISCVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.CommissionAmount, 0)      + ISNULL(ctvCost.COMVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.ServiceFeeAmount, 0)      + ISNULL(ctvCost.SFVatAmount, 0)          " &
    '                        " 							+ ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate               " &
    '                        " 							ELSE 0 END))" &
    '                        " 		/(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax						" &
    '                        " FROM dbo.CommercialTransactions WITH (NOLOCK)    " &
    '                        " 	INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID   " &
    '                        " 	LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1   " &
    '                        " 	LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'    " &
    '                        " 	INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID    " &
    '                        " 	RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)     " &
    '                        " 		INNER JOIN dbo.Documents WITH (NOLOCK)      " &
    '                        " 			INNER JOIN dbo.TFEntities WITH (NOLOCK)      " &
    '                        " 				LEFT JOIN dbo.TFEntityTags        " &
    '                        " 					LEFT JOIN dbo.Tags         ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID       " &
    '                        " 				ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)      " &
    '                        " 			ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id      " &
    '                        " 		ON dbo.DocTypes.Id = dbo.Documents.DocTypesID     " &
    '                        " 	ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id   " &
    '                        " WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'       " &
    '                        " 	  AND (dbo.Documents.IsCancellationDocument = 0)    " &
    '                        " 	  AND (dbo.Documents.DocStatusID = 41)   " &
    '                        " 	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)   " &
    '                        " 	  AND dbo.CommercialTransactionValues.Id IS NOT NULL     " &
    '                        " 	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  " &
    '                        " GROUP BY ISNULL(dbo.Tags.Description, '')   " &
    '                        " 		 , dbo.TFEntities.Code  " &
    '                        " 		 , dbo.TFEntities.Name   " &
    '                        " ORDER BY ISNULL(dbo.Tags.Description, '')   " &
    '                        " 		,dbo.TFEntities.Code "
    '    End Function
    '    Public Function E10_ProfitPerClientGroupWithExtra(ByVal TagGroup As Integer, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = FromDate
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = ToDate
    '        sqlComm.CommandText = ProfitPerClientGroupStatement(True)
    '        Return sqlComm
    '    End Function
    '    Public Function E11_ProfitPerOPSGroupWithPY(ByVal TagGroup As Integer, ByVal DateFrom As Date, ByVal DateTo As Date, ByVal FromYTD As Date, ByVal ToYTD As Date, ByVal FromPY As Date, ByVal ToPY As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = DateFrom
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = DateTo
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = ToYTD
    '        sqlComm.Parameters.Add("@FromPY", SqlDbType.Date).Value = FromPY
    '        sqlComm.Parameters.Add("@ToPY", SqlDbType.Date).Value = ToPY
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = " USE TravelForceCosmos  " &
    '                            "   If(OBJECT_ID('tempdb..#TempTablePY') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTablePY   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableCurr   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableBudget') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableBudget   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableClients   " &
    '                            "   End  " &
    '                            "     SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTableCurr   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL      " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            "   SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTablePY   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromPY AND @ToPY)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL      " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            "   SELECT dbo.TFEntities.Code AS Client      " &
    '                            "   	 , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN   " &
    '                            "   	                             (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "   	                             + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "   	                             * dbo.CommercialTransactionValues.Rate                          " &
    '                            "   	                             ELSE 0 END))                       AS Sales          " &
    '                            "   	 , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                    (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    " &
    '                            "                                    + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    " &
    '                            "                                    * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                    ELSE 0 END)      " &
    '                            "                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                          (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                          + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                          + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                          + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                          + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                          + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                          + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                          + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                          + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                          * ctvCost.Rate                          " &
    '                            "                                          ELSE 0 END)) AS Profit             " &
    '                            "        , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax         " &
    '                            "        , CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                      ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)   " &
    '                            "                                      + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))   " &
    '                            "                                      * dbo.CommercialTransactionValues.Rate                          " &
    '                            "                                      ELSE 0 END)      " &
    '                            "                                      + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                    " &
    '                            "                                            (ISNULL(ctvCost.FaceValue, 0)                 " &
    '                            "                                            + ISNULL(ctvCost.FVVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.FaceValueExtra, 0)           " &
    '                            "                                            + ISNULL(ctvCost.FVXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.Taxes, 0)                    " &
    '                            "                                            + ISNULL(ctvCost.TAXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.TaxesExtra, 0)               " &
    '                            "                                            + ISNULL(ctvCost.TAXXVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.DiscountAmount, 0)           " &
    '                            "                                            + ISNULL(ctvCost.DISCVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CommissionAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.COMVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.ServiceFeeAmount, 0)         " &
    '                            "                                            + ISNULL(ctvCost.SFVatAmount, 0)                     " &
    '                            "                                            + ISNULL(ctvCost.CancellationFeeAmount, 0)    " &
    '                            "                                            + ISNULL(ctvCost.CFVatAmount, 0))    " &
    '                            "                                            * ctvCost.Rate                          " &
    '                            "                                            ELSE 0 END))      " &
    '                            "                                            /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))) AS ProfitPerPax      " &
    '                            "   INTO #TempTableYTD   " &
    '                            "   FROM dbo.CommercialTransactions WITH (NOLOCK)         " &
    '                            "   INNER JOIN dbo.CommercialTransactionValues    " &
    '                            "   ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        " &
    '                            "   LEFT JOIN dbo.CommercialTransactionValues ctvCost    " &
    '                            "   ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID    " &
    '                            "      AND ctvCost.IsCost=1        " &
    '                            "   INNER JOIN dbo.DocumentItems WITH (NOLOCK)    " &
    '                            "   ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID         " &
    '                            "   RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)           " &
    '                            "   	INNER JOIN dbo.Documents WITH (NOLOCK)             " &
    '                            "   		INNER JOIN dbo.TFEntities WITH (NOLOCK)              " &
    '                            "   		ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id            " &
    '                            "   	ON dbo.DocTypes.Id = dbo.Documents.DocTypesID          " &
    '                            "   ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id       " &
    '                            "   WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'              " &
    '                            "   	  AND (dbo.Documents.IsCancellationDocument = 0)           " &
    '                            "   	  AND (dbo.Documents.DocStatusID = 41)          " &
    '                            "         AND (Documents.DocTypesID NOT IN (74, 75))" &
    '                            "   	  AND (dbo.Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)          " &
    '                            "   	  AND dbo.CommercialTransactionValues.Id IS NOT NULL            " &
    '                            "   	  AND dbo.DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL     " &
    '                            "   GROUP BY dbo.TFEntities.Code   " &
    '                            "   ORDER BY dbo.TFEntities.Code   " &
    '                            " SELECT TFEntities.Code " &
    '                            " INTO #TempTableClients " &
    '                            " FROM TFEntities " &
    '                            " LEFT JOIN #TempTableCurr ON #TempTableCurr.Client=TFEntities.Code " &
    '                            " LEFT JOIN #TempTableYTD ON #TempTableYTD.Client=TFEntities.Code " &
    '                            " LEFT JOIN #TempTablePY ON #TempTablePY.Client=TFEntities.Code " &
    '                            " WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL OR #TempTablePY.Client IS NOT NULL) " &
    '                            " 	  AND (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0 " &
    '                            " 	  OR   #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0 " &
    '                            " 	  OR   #TempTablePY.Sales<>0 OR #TempTablePY.Profit <> 0 OR #TempTablePY.Pax <>0) " &
    '                            "   SELECT  " &
    '                            "    ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED') AS GroupName        " &
    '                            "   	 , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name) AS Client      " &
    '                            "	     , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  " &
    '                            "        , COALESCE(SUM(#TempTableCurr.Profit)/NULLIF(SUM(#TempTableCurr.Pax),0),0) AS ProfitPerPax  " &
    '                            "	     , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD   " &
    '                            "        , COALESCE(SUM(#TempTableYTD.Profit)/NULLIF(SUM(#TempTableYTD.Pax),0),0) AS ProfitPerPaxYTD   " &
    '                            "        , COALESCE(SUM(#TempTablePY.Sales),0) AS SalesPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Profit),0) AS ProfitPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Pax),0) AS PaxPY                                                 " &
    '                            "        , COALESCE(SUM(#TempTablePY.Profit)/NULLIF(SUM(#TempTablePY.Pax),0),0) AS ProfitPerPaxPY                                                 " &
    '                            "   FROM #TempTableClients         " &
    '                            " 	LEFT JOIN TFEntities " &
    '                            " 	ON #TempTableClients.Code = TFEntities.Code " &
    '                            "   		LEFT JOIN dbo.TFEntityTags                 " &
    '                            "   			LEFT JOIN dbo.Tags    " &
    '                            "   			ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID               " &
    '                            "   		ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID    " &
    '                            "   		   AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)              " &
    '                            "   	LEFT JOIN dbo.TFEntityTags TFEntityTagsClientGroup                 " &
    '                            "   		LEFT JOIN dbo.Tags TagClientGroup    " &
    '                            "   		ON TagClientGroup.TagGroupID=146    " &
    '                            "   			AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID        " &
    '                            "   	ON dbo.TFEntities.Id = TFEntityTagsClientGroup.TFEntityID    " &
    '                            "   		AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=146 AND dbo.Tags.Id=TFEntityTagsClientGroup.TagID)                 " &
    '                            "   LEFT JOIN #TempTableCurr " &
    '                            " 	ON #TempTableCurr.Client=TFEntities.Code   " &
    '                            "   LEFT JOIN #TempTablePY   " &
    '                            "   	ON #TempTablePY.Client = TFEntities.Code   " &
    '                            "   LEFT JOIN #TempTableYTD   " &
    '                            "   	ON #TempTableYTD.Client = TFEntities.Code	      " &
    '                            "   GROUP BY ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED') " &
    '                            "            , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)       " &
    '                            "   ORDER BY ISNULL(dbo.Tags.Description, '00 - UNCLASSIFIED')      " &
    '                            "            , ISNULL(TagClientGroup.Description, dbo.TFEntities.Code + '/' + dbo.TFEntities.Name)     " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableCurr   " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTablePY') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTablePY   " &
    '                            "   End   " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD  " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableBudget') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableYTD  " &
    '                            "   End " &
    '                            "   If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)   " &
    '                            "   Begin   " &
    '                            "   Drop Table #TempTableClients    " &
    '                            "   End "

    '        Return sqlComm
    '    End Function
    '    Public Function E12_ProfitPerOPSGroupWithBudgetComparison(ByVal TagGroup As Integer, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = TagGroup
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.E12_FromCurr
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.E12_ToCurr
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToYTD", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@FromPYTD", SqlDbType.Date).Value = mReport.E12_FromPYTD
    '        sqlComm.Parameters.Add("@ToPYTD", SqlDbType.Date).Value = mReport.E12_ToPYTD
    '        sqlComm.Parameters.Add("@FromPYCurr", SqlDbType.Date).Value = mReport.E12_FromPYCurr
    '        sqlComm.Parameters.Add("@ToPYCurr", SqlDbType.Date).Value = mReport.E12_ToPYCurr
    '        sqlComm.Parameters.Add("@CurrYear", SqlDbType.Int).Value = mReport.ReportYear
    '        sqlComm.Parameters.Add("@FromMonth", SqlDbType.Int).Value = 1
    '        sqlComm.Parameters.Add("@ToMonth", SqlDbType.Int).Value = mReport.ReportMonth
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTablePYCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTablePYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableBudgetCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableBudgetYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWYTD  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWPYCurr  
    'End  
    'If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  
    'Begin  
    'Drop Table #TempTableIWPYtd  
    'End  

    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  
    'Begin  
    'Drop Table #TempTableClients  
    'End

    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWCurr  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '        ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '        ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    'GROUP BY CommercialTransactionValueID  

    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWYTD  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE  SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )  
    '        )
    'GROUP BY CommercialTransactionValueID  
    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWPYCurr  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )
    '        )
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        )
    '        ) 		  



    'GROUP BY CommercialTransactionValueID  
    'SELECT  CommercialTransactionValueID, SUM(Amount) AS IWAmount  
    'INTO #TempTableIWPYtd  
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis  
    'WHERE (ServiceFeeTypeID IN (1,3,4,5,6) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        ) 
    '        ) 
    '        OR
    '(ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (  

    'SELECT DISTINCT CommercialTransactionValues.Id   

    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '    RIGHT JOIN travelForceCosmos.dbo.ServiceFeeAnalysis  
    '    ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '        AND ServiceFeeAnalysis.Id IS NOT NULL  
    '        AND CommercialTransactionValues.IsCost=0  
    '        ) 
    '        ) 
    'GROUP BY CommercialTransactionValueID  

    'SELECT   ISNULL(Code, tfbOtherCode) AS Client  
    '        ,ISNULL(Name, tfbOtherName) AS ClientName
    '        ,tfbTFEntityId
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  
    '        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  
    '        ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  
    'INTO #TempTableBudgetCurr  
    'FROM [ATPIData].[dbo].[TFClientBudget]  
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities  
    'ON tfbTFEntityId = TFEntities.Id  
    'WHERE tfbYear = @CurrYear AND tfbMonth = @ToMonth  
    'GROUP BY ISNULL(Code, tfbOtherCode),ISNULL(Name, tfbOtherName),tfbTFEntityId  
    'ORDER BY ISNULL(Code, tfbOtherCode)  

    'SELECT   ISNULL(Code, tfbOtherCode) AS Client  
    '        ,ISNULL(Name, tfbOtherName) AS ClientName
    '        ,tfbTFEntityId
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbTurnover,0))) AS Sales  
    '        ,CONVERT(DECIMAL(18,0),SUM(ISNULL(tfbMargin  ,0)))   AS Profit  
    '        ,SUM(ISNULL(tfbTransactions,0)) AS Pax  
    '        ,SUM(ISNULL(tfbMargin,0))/(NULLIF(SUM(ISNULL(tfbTransactions, 0) ), 0)) AS ProfitPerPax  
    'INTO #TempTableBudgetYTD  
    'FROM [ATPIData].[dbo].[TFClientBudget]  
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities  
    'ON tfbTFEntityId = TFEntities.Id  
    'WHERE tfbYear = @CurrYear AND tfbMonth BETWEEN @FromMonth AND @ToMonth  
    'GROUP BY ISNULL(Code, tfbOtherCode),ISNULL(Name, tfbOtherName),tfbTFEntityId  
    'ORDER BY ISNULL(Code, tfbOtherCode)  

    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWCurr.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTableCurr  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWCurr  
    '        ON #TempTableIWCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL 

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWYTD.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTableYTD  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWYTD  
    '        ON #TempTableIWYTD.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL  

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYCurr.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTablePYCurr  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWPYCurr  
    '        ON #TempTableIWPYCurr.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYCurr AND @ToPYCurr)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  
    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    '    SELECT TFEntities.Code AS Client  
    '        , CONVERT(DECIMAL(18,0), SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0)   
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END))                       AS Sales  
    '        , CONVERT(DECIMAL(18,0),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                    (ISNULL(CommercialTransactionValues.FaceValue, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.Taxes, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TaxesExtra, 0)  
    '                                    + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DiscountAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CommissionAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.COMVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) - ISNULL(#TempTableIWPYtd.IWAmount,0)  
    '                                    + ISNULL(CommercialTransactionValues.SFVatAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)  
    '                                    + ISNULL(CommercialTransactionValues.CFVatAmount, 0))  
    '                                    * CommercialTransactionValues.Rate  
    '                                    ELSE 0 END)  
    '                                    + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN  
    '                                        (ISNULL(ctvCost.FaceValue, 0)  
    '                                        + ISNULL(ctvCost.FVVatAmount, 0)  
    '                                        + ISNULL(ctvCost.FaceValueExtra, 0)  
    '                                        + ISNULL(ctvCost.FVXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.Taxes, 0)  
    '                                        + ISNULL(ctvCost.TAXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.TaxesExtra, 0)  
    '                                        + ISNULL(ctvCost.TAXXVatAmount, 0)  
    '                                        + ISNULL(ctvCost.DiscountAmount, 0)  
    '                                        + ISNULL(ctvCost.DISCVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CommissionAmount, 0)  
    '                                        + ISNULL(ctvCost.COMVatAmount, 0)  
    '                                        + ISNULL(ctvCost.ServiceFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.SFVatAmount, 0)  
    '                                        + ISNULL(ctvCost.CancellationFeeAmount, 0)  
    '                                        + ISNULL(ctvCost.CFVatAmount, 0))  
    '                                        * ctvCost.Rate  
    '                                        ELSE 0 END)) AS Profit  
    '        , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS Pax  
    '    INTO #TempTablePYTD  
    '    FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues  
    '        LEFT JOIN #TempTableIWPYtd  
    '        ON #TempTableIWPYtd.CommercialTransactionValueID = CommercialTransactionValues.Id  
    '    ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID  
    '    LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost  
    '    ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID  
    '    AND ctvCost.IsCost=1  
    '    INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)  
    '    ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID  
    '    RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)  
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)  
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)  
    '        ON Documents.CounterPartyID = TFEntities.Id  
    '    ON DocTypes.Id = Documents.DocTypesID  
    '    ON DocumentItems.DocumentsID = Documents.Id  
    '    WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'  
    '        AND (Documents.IsCancellationDocument = 0)  
    '        AND (Documents.DocStatusID = 41)  
    '        AND (Documents.DocTypesID NOT IN (74, 75))  
    '        AND (Documents.IssueDate BETWEEN  @FromPYTD AND @ToPYTD)  
    '        AND CommercialTransactionValues.Id IS NOT NULL  
    '        AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL  

    '    GROUP BY TFEntities.Code  
    '    ORDER BY TFEntities.Code  
    'SELECT TFEntities.Code  
    'INTO #TempTableClients  
    'FROM TravelForceCosmos.dbo.TFEntities  
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code  
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code  
    'LEFT JOIN #TempTablePYCurr     ON #TempTablePYCurr.Client     = TFEntities.Code  
    'LEFT JOIN #TempTablePYTD       ON #TempTablePYTD.Client       = TFEntities.Code  
    'LEFT JOIN #TempTableBudgetCurr ON #TempTableBudgetCurr.Client = TFEntities.Code  
    'LEFT JOIN #TempTableBudgetYTD  ON #TempTableBudgetYTD.Client  = TFEntities.Code  
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL   
    '    OR #TempTablePYTD.Client IS NOT NULL OR #TempTablePYCurr.Client IS NOT NULL   
    '    OR #TempTableBudgetCurr.Client IS NOT NULL OR #TempTableBudgetYTD.Client IS NOT NULL)  
    '    AND 
    '    (#TempTableCurr.Sales<>0 OR #TempTableCurr.Profit <> 0 OR #TempTableCurr.Pax <>0  
    '    OR #TempTableYTD.Sales<>0 OR #TempTableYTD.Profit <> 0 OR #TempTableYTD.Pax <>0  
    '    OR #TempTablePYTD.Sales<>0 OR #TempTablePYTD.Profit <> 0 OR #TempTablePYTD.Pax <>0  
    '    OR #TempTablePYCurr.Sales<>0 OR #TempTablePYCurr.Profit <> 0 OR #TempTablePYCurr.Pax <>0  
    '    OR #TempTableBudgetCurr.Sales<>0 OR #TempTableBudgetCurr.Profit <> 0 OR #TempTableBudgetCurr.Pax <> 0  
    '    OR #TempTableBudgetYTD.Sales<>0 OR #TempTableBudgetYTD.Profit <> 0 OR #TempTableBudgetYTD.Pax <> 0)  

    '    SELECT  
    '    ISNULL(Tags.Description, '00 - UNCLASSIFIED') AS GroupName  
    '        , ISNULL(TagClientGroup.Description, TFEntities.Code + '/' + TFEntities.Name) AS Client  
    '        , COALESCE(SUM(#TempTableCurr.Sales),0) AS Sales  
    '        , COALESCE(SUM(#TempTableCurr.Profit),0) AS Profit  
    '        , COALESCE(SUM(#TempTableCurr.Pax),0) AS Pax  
    '        , 0 AS ProfitPerPax  
    '        , COALESCE(SUM(#TempTableYTD.Sales),0) AS SalesYTD  
    '        , COALESCE(SUM(#TempTableYTD.Profit),0) AS ProfitYTD  
    '        , COALESCE(SUM(#TempTableYTD.Pax),0) AS PaxYTD  
    '        , 0 AS ProfitPerPaxYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Sales),0) AS SalesPYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Profit),0) AS ProfitPYTD  
    '        , COALESCE(SUM(#TempTablePYTD.Pax),0) AS PaxPYTD  
    '        , 0 AS ProfitPerPaxPYTD  
    '        , COALESCE(SUM(#TempTablePYCurr.Sales),0) AS SalesPYCurr  
    '        , COALESCE(SUM(#TempTablePYCurr.Profit),0) AS ProfitPYCurr  
    '        , COALESCE(SUM(#TempTablePYCurr.Pax),0) AS PaxPYCurr  
    '        , 0 AS ProfitPerPaxPYCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  
    '        , 0 AS ProfitPerPaxBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  
    '        , 0 AS ProfitPerPaxBudgetYTD  
    'FROM #TempTableClients  
    'LEFT JOIN TFEntities  
    'ON #TempTableClients.Code = TFEntities.Code  
    '        LEFT JOIN TravelForceCosmos.dbo.TFEntityTags  
    '            LEFT JOIN TravelForceCosmos.dbo.Tags  
    '            ON Tags.TagGroupID=@TagGroup AND Tags.Id=dbo.TFEntityTags.TagID  
    '        ON TFEntities.Id = TFEntityTags.TFEntityID  
    '            AND TFEntityTags.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WHERE Tags.TagGroupID=@TagGroup AND Tags.Id=dbo.TFEntityTags.TagID)  
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup  
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup  
    '        ON TagClientGroup.TagGroupID=146  
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID  
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID  
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)  
    '    LEFT JOIN #TempTableCurr  
    '    ON #TempTableCurr.Client=TFEntities.Code  
    '    LEFT JOIN #TempTablePYTD  
    '    ON #TempTablePYTD.Client = TFEntities.Code  
    '    LEFT JOIN #TempTableYTD  
    '    ON #TempTableYTD.Client = TFEntities.Code	  
    '    LEFT JOIN #TempTableBudgetCurr  
    '    ON #TempTableBudgetCurr.Client = TFEntities.Code  
    '    LEFT JOIN #TempTableBudgetYTD  
    '    ON #TempTableBudgetYTD.Client = TFEntities.Code  
    '    LEFT JOIN #TempTablePYCurr  
    '    ON #TempTablePYCurr.Client = TFEntities.Code  
    '    GROUP BY ISNULL(Tags.Description, '00 - UNCLASSIFIED')  
    '            , ISNULL(TagClientGroup.Description, TFEntities.Code + '/' + TFEntities.Name)  

    'UNION

    'SELECT '99 - OTHER' AS GroupName  
    '     , #TempTableBudgetCurr.Client + '/' + #TempTableBudgetCurr.ClientName AS Client
    '        , 0 AS Sales  
    '        , 0 AS Profit  
    '        , 0 AS Pax  
    '        , 0 AS ProfitPerPax  
    '        , 0 AS SalesYTD  
    '        , 0 AS ProfitYTD  
    '        , 0 AS PaxYTD  
    '        , 0 AS ProfitPerPaxYTD  
    '        , 0 AS SalesPYTD  
    '        , 0 AS ProfitPYTD  
    '        , 0 AS PaxPYTD  
    '        , 0 AS ProfitPerPaxPYTD  
    '        , 0 AS SalesPYCurr  
    '        , 0 AS ProfitPYCurr  
    '        , 0 AS PaxPYCurr  
    '        , 0 AS ProfitPerPaxPYCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Sales),0) AS SalesBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Profit),0) AS ProfitBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetCurr.Pax),0) AS PaxBudgetCurr  
    '        , 0 AS ProfitPerPaxBudgetCurr  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Sales),0) AS SalesBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Profit),0) AS ProfitBudgetYTD  
    '        , COALESCE(SUM(#TempTableBudgetYTD.Pax),0) AS PaxBudgetYTD  
    '        , 0 AS ProfitPerPaxBudgetYTD  

    'FROM #TempTableBudgetCurr
    'LEFT JOIN #TempTableBudgetYTD
    'ON  #TempTableBudgetCurr.Client     = #TempTableBudgetYTD.Client
    'AND #TempTableBudgetCurr.ClientName = #TempTableBudgetYTD.ClientName
    'WHERE #TempTableBudgetCurr.tfbTFEntityId IS NULL
    'GROUP BY #TempTableBudgetCurr.Client + '/' + #TempTableBudgetCurr.ClientName
    '    ORDER BY GroupName, Client

    '    If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTablePYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTablePYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTablePYCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTablePYCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableBudgetCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableBudgetCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableBudgetYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableBudgetYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableClients  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWYTD') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWYTD  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWPYCurr') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWPYCurr  
    '    End  
    '    If(OBJECT_ID('tempdb..#TempTableIWPYtd') Is Not Null)  
    '    Begin  
    '    Drop Table #TempTableIWPYtd  
    '    End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E13_TicketAnalysis(ByVal UninvoicedOnly As Boolean, ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromIssue", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToIssue", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromDep", SqlDbType.Date).Value = mReport.E12_FromYTD
    '        sqlComm.Parameters.Add("@ToDep", SqlDbType.Date).Value = mReport.E12_ToYTD
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "  USE TravelForceCosmos  
    '          SELECT dbo.TFEntities.Code AS ClientCode
    '        , dbo.TFEntities.Name AS ClientName
    '        , dbo.LookupTable.Name AS GDS
    '        , dbo.CommercialTransactions.[CreatorSalesmanString]
    '        , dbo.CommercialTransactions.[CreatorPCC]
    '        , dbo.CommercialTransactions.[IssueSalesmanString]
    '        , dbo.CommercialTransactions.[IssuePCC]
    '        , dbo.CommercialTransactions.[CustomDescription2] AS PNR
    '        , dbo.Airlines.IATAAccountingPrefix AS AirlineCode 
    '        , dbo.Airlines.IATACode AS AirlineAbbreviation
    '        , dbo.CommercialTransactions.[CustomDescription1] AS TicketNumber
    '        , dbo.AirTickets.IssueDate AS TicketIssueDate
    '        , ISNULL(dbo.AirTickets.NetRemitCode, '') AS TourCode
    '        , dbo.CommercialTransactions.[CustomDescription3] AS Passengers
    '        , dbo.CommercialTransactions.[CustomDescription4] AS Routing
    '        , dbo.CommercialTransactions.[FromDate] AS [Departure Date]
    '        , ISNULL(CPVBookedBy.Value, '') AS BookedBy      
    '        , ISNULL(CPVCostCentre.Value, '') AS CostCentre
    '        , ISNULL(dbo.TFEntityDepartments.Name, '') AS Vessel
    '        , ISNULL(CPVTrId.Value, '') AS TRID
    '        , ISNULL(dbo.Documents.IssueDate, '') AS InvoiceDate
    '        , ISNULL(dbo.DocTypes.Code, '') AS InvType
    '        , ISNULL(dbo.DocTypes.Series, '') AS InvSeries
    '        , ISNULL(dbo.Documents.Number, '') AS InvoiceNumber
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Fare          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Taxes 
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    
    '                                 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS CancFee 
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Discount          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       
    '                                 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Commission          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         
    '                                 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS ServiceFee          
    '        , CONVERT(DECIMAL(18,2), (   
    '                                 ( ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)                 
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                    
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)               
    '                                 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)           
    '                                 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)       
    '                                 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)         
    '                                 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                     
    '                                 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0)    
    '                                 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0))    
    '                                 * dbo.CommercialTransactionValues.Rate                          
    '                                 ))                       AS Sales          
    '           , (ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax
    'FROM dbo.CommercialTransactions WITH (NOLOCK)         
    ' INNER JOIN dbo.CommercialTransactionValues    
    ' ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID        
    '    AND dbo.CommercialTransactionValues.IsCost=0
    'INNER JOIN dbo.TFEntities WITH (NOLOCK)              
    'ON dbo.[CommercialTransactionValues].[CommercialEntityID] = dbo.TFEntities.Id
    'LEFT JOIN dbo.TFEntityDepartments
    'ON dbo.[CommercialTransactionValues].[CommercialEntityDepartmentID]= dbo.TFEntityDepartments.Id            
    'LEFT JOIN dbo.AirTickets
    '    ON dbo.CommercialTransactions.CustomDescription1 = dbo.AirTickets.DocumentNr
    'LEFT JOIN dbo.Airlines
    '    ON dbo.AirTickets.TicketingAirlineID = dbo.Airlines.Id		     
    'LEFT JOIN dbo.CustomPropertyValues CPVBookedBy
    '    ON dbo.AirTickets.BookFileId = CPVBookedBy.BookFileId 
    '        AND CPVBookedBy.CustomPropertyID=1 
    '        AND CPVBookedBy.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.CustomPropertyValues CPVCostCentre
    '    ON dbo.CommercialTransactionValues.CommercialTransactionID = CPVCostCentre.CTID 
    '        AND CPVCostCentre.CustomPropertyID=5 
    '        AND CPVCostCentre.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.CustomPropertyValues CPVTrId
    '    ON dbo.AirTickets.BookFileId = CPVTrId.BookFileId 
    '        AND CPVTrId.CustomPropertyID=14 
    '        AND CPVCostCentre.TFEntityId=dbo.[CommercialTransactionValues].[CommercialEntityID]
    'LEFT JOIN dbo.DocumentItems 
    '    ON dbo.DocumentItems.[CommercialTransactionValueID]= dbo.CommercialTransactionValues.Id  
    'LEFT JOIN dbo.Documents
    '    ON dbo.DocumentItems.DocumentsId = dbo.Documents.Id	    
    'LEFT JOIN dbo.DocTypes
    '    ON dbo.Documents.DocTypesId = dbo.DocTypes.Id
    'LEFT JOIN dbo.LookupTable
    '    ON dbo.AirTickets.GDSSystemId = dbo.LookupTable.Id
    ' WHERE    dbo.TFEntities.Code = @ClientCode  
    '          -- dbo.TFEntities.ID IN (SELECT TFEntityID FROM dbo.TFEntitytags WHERE TagId IN (154,155) )"
    '        If UninvoicedOnly Then
    '            sqlComm.CommandText &= " AND dbo.DocumentItems.Id IS NULL "
    '        End If
    '        sqlComm.CommandText &= "AND StatusID <>738 -- Not Cancelled
    '      AND dbo.CommercialTransactionValues.Id IS NOT NULL  
    '      AND dbo.AirTickets.IssueDate BETWEEN @FromIssue AND @ToIssue          
    'ORDER BY dbo.TFEntities.Code, dbo.TFEntities.Name, dbo.CommercialTransactions.[FromDate], dbo.AirTickets.IssueDate, dbo.CommercialTransactions.[CustomDescription2], dbo.CommercialTransactions.[CustomDescription1]   "
    '        Return sqlComm
    '    End Function
    '    Public Function E16_DailyProfitReport(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '        sqlComm.CommandTimeout = 400
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    'Begin
    'Drop Table #TempTableYTD
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWytd5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    'Begin
    'Drop Table #TempIWytd10
    'End
    'If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    'Begin
    'Drop Table #TempUninvoiced
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempUninvoiced -------------------------------------------------------------------------------------------------------------------
    ' SELECT TFEntities.Code AS Client
    '      , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate
    '            ELSE 0 END )) AS NetPayableUninvoiced
    '      , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS PaxUninvoiced
    ' INTO #TempUninvoiced
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID AND CommercialTransactionValues.IsCost = 0
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON CommercialTransactionValues.CommercialEntityID = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.AirTicketTransactions WITH (NOLOCK)
    'ON CommercialTransactions.Id = AirTicketTransactions.CommercialTransactionID
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND DocumentItems.Id IS NULL
    '      AND (CommercialTransactions.TransactionDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND StatusId = 331
    '      AND IsReversed = 0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    ' GROUP BY TFEntities.Code 
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0      
    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0      
    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempIWytd5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWytd10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableYTD --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableYTD
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWytd5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWytd10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code
    'LEFT JOIN #TempUninvoiced      ON #TempUninvoiced.Client      = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0
    '	OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0
    '    OR #TempUninvoiced.PaxUninvoiced >0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 1 AS Tots
    '      , ISNULL(TagClientGroup.Description, '') AS ClientGroupDescription
    '      , CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END AS ClientCode
    '      , ISNULL(TagClientGroup.Description, TFEntities.Name) AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' GROUP BY ISNULL(TagClientGroup.Description, ''),  CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END, ISNULL(TagClientGroup.Description, TFEntities.Name)

    'UNION

    '  SELECT 2 AS Tots
    '      , TagClientGroup.Description AS ClientGroupDescription
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' WHERE TagClientGroup.Description IS NOT NULL

    ' GROUP BY  TagClientGroup.Description, TFEntities.Code, TFEntities.Name

    ' ORDER BY Tots, Profit DESC
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    ' Begin
    ' Drop Table #TempTableYTD
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr10
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    ' Begin
    ' Drop Table #TempUninvoiced
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd10
    ' End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E15_DailyProfitReportWithoutRINVA(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@FromYTD", SqlDbType.Date).Value = DateSerial(mReport.Date1From.Year, 1, 1)
    '        sqlComm.CommandTimeout = 400
    '        sqlComm.CommandText = "
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    'Begin
    'Drop Table #TempTableYTD
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWytd5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    'Begin
    'Drop Table #TempIWytd10
    'End
    'If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    'Begin
    'Drop Table #TempUninvoiced
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempUninvoiced -------------------------------------------------------------------------------------------------------------------
    ' SELECT TFEntities.Code AS Client
    '      , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate
    '            ELSE 0 END )) AS NetPayableUninvoiced
    '      , SUM(ISNULL(CommercialTransactions.Pax, 0)) AS PaxUninvoiced
    ' INTO #TempUninvoiced
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID AND CommercialTransactionValues.IsCost = 0
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON CommercialTransactionValues.CommercialEntityID = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.AirTicketTransactions WITH (NOLOCK)
    'ON CommercialTransactions.Id = AirTicketTransactions.CommercialTransactionID
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND DocumentItems.Id IS NULL
    '      AND (CommercialTransactions.TransactionDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND StatusId = 331
    '      AND IsReversed = 0
    ' GROUP BY TFEntities.Code 
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempIWytd5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWytd10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWytd10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    ' RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN(154,155))
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableYTD --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '     ,  CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID=1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END)
    '         AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID<>1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    ' INTO #TempTableYTD
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWytd5to9and11 IWTable WITH (NOLOCK)
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWytd10 IW10Table WITH (NOLOCK)
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE SUBSTRING(TFEntities.Code,1,1) <= '0'
    '      AND (SELECT COUNT(*) FROM AmadeusReports.dbo.TFReportExclude WHERE TFReportExclude.ReportNumber = 15 AND TFReportExclude.ClientCode = TFEntities.Code)=0
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75, 109)) -- removed 134 (8/10/2022)
    '      AND (Documents.IssueDate BETWEEN  @FromYTD AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)  --AND DocTypes.AccGeneratorsID IS NOT NULL

    ' GROUP BY TFEntities.Code
    ' ORDER BY TFEntities.Code
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'LEFT JOIN #TempTableYTD        ON #TempTableYTD.Client        = TFEntities.Code
    'LEFT JOIN #TempUninvoiced      ON #TempUninvoiced.Client      = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL OR #TempTableYTD.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0
    '	OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0
    '    OR #TempUninvoiced.PaxUninvoiced >0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 1 AS Tots
    '      , ISNULL(TagClientGroup.Description, '') AS ClientGroupDescription
    '      , CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END AS ClientCode
    '      , ISNULL(TagClientGroup.Description, TFEntities.Name) AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' GROUP BY ISNULL(TagClientGroup.Description, ''),  CASE WHEN TagClientGroup.Description IS NOT NULL THEN '' ELSE TFEntities.Code END, ISNULL(TagClientGroup.Description, TFEntities.Name)

    'UNION

    '  SELECT 2 AS Tots
    '      , TagClientGroup.Description AS ClientGroupDescription
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0),0),0)) AS ProfitPerPax
    ' -- Year To Date ----
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) AS NetPayableYTDAir
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) AS NetBuyYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) AS IW05YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) AS IW06YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) AS IW07YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) AS IW08YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) AS IW09YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) AS IW11YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) AS IW10YTDAIR
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) AS IWYTDAIR
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) AS ProfitYTDAir
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) AS PaxYTDAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0),0),0)) AS ProfitPerPaxYTDAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW05Services),0) AS IW05YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW06Services),0) AS IW06YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW07Services),0) AS IW07YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW08Services),0) AS IW08YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW09Services),0) AS IW09YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW11Services),0) AS IW11YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IW10Services),0) AS IW10YTDServices
    '      , COALESCE(SUM(#TempTableYTD.IWServices),0) AS IWYTDServices
    '      , COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTDServices
    '      , COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTDServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTDServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) AS NetPayableYTD
    '      , COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) + COALESCE(SUM(#TempTableYTD.NetBuyServices),0) AS NetBuyYTD
    '      , COALESCE(SUM(#TempTableYTD.IW05AIR),0) + COALESCE(SUM(#TempTableYTD.IW05Services),0)  AS IW05YTD
    '      , COALESCE(SUM(#TempTableYTD.IW06AIR),0) + COALESCE(SUM(#TempTableYTD.IW06Services),0)  AS IW06YTD
    '      , COALESCE(SUM(#TempTableYTD.IW07AIR),0) + COALESCE(SUM(#TempTableYTD.IW07Services),0)  AS IW07YTD
    '      , COALESCE(SUM(#TempTableYTD.IW08AIR),0) + COALESCE(SUM(#TempTableYTD.IW08Services),0)  AS IW08YTD
    '      , COALESCE(SUM(#TempTableYTD.IW09AIR),0) + COALESCE(SUM(#TempTableYTD.IW09Services),0)  AS IW09YTD
    '      , COALESCE(SUM(#TempTableYTD.IW11AIR),0) + COALESCE(SUM(#TempTableYTD.IW11Services),0)  AS IW11YTD
    '      , COALESCE(SUM(#TempTableYTD.IW10AIR),0) + COALESCE(SUM(#TempTableYTD.IW10Services),0)  AS IW10YTD
    '      , COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.IWServices),0)  AS IWYTD
    '      , COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0) AS ProfitYTD
    '      , COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0) AS PaxYTD
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) AS ProfitPerPaxYTD
    '      , COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS PaxUninvoiced
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableYTD.NetPayableAir),0) - COALESCE(SUM(#TempTableYTD.NetBuyAIR),0) - COALESCE(SUM(#TempTableYTD.IWAIR),0) + COALESCE(SUM(#TempTableYTD.NetPayableServices),0) - COALESCE(SUM(#TempTableYTD.NetBuyServices),0) - COALESCE(SUM(#TempTableYTD.IWServices),0))/NULLIF(COALESCE(SUM(#TempTableYTD.PaxAir),0) + COALESCE(SUM(#TempTableYTD.PaxServices),0),0),0)) * COALESCE(SUM(#TempUninvoiced.PaxUninvoiced),0) AS ProfitUninvoicedPax
    '	  , COALESCE(SUM(#TempUninvoiced.NetPayableUninvoiced), 0) AS NetPayableUninvoiced
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    '    LEFT JOIN TravelForceCosmos.dbo.TFEntityTags TFEntityTagsClientGroup WITH (NOLOCK)
    '        LEFT JOIN TravelForceCosmos.dbo.Tags TagClientGroup WITH (NOLOCK)
    '        ON TagClientGroup.TagGroupID=146
    '            AND TagClientGroup.Id=TFEntityTagsClientGroup.TagID
    '    ON TFEntities.Id = TFEntityTagsClientGroup.TFEntityID
    '        AND TFEntityTagsClientGroup.TagID IN (SELECT Id FROM TravelForceCosmos.dbo.Tags WITH (NOLOCK) WHERE Tags.TagGroupID=146 AND Tags.Id=TFEntityTagsClientGroup.TagID)
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' LEFT JOIN #TempTableYTD
    '    ON #TempTableYTD.Client = TFEntities.Code	
    ' LEFT JOIN #TempUninvoiced
    '    ON #TempUninvoiced.Client = TFEntities.Code	
    ' WHERE TagClientGroup.Description IS NOT NULL

    ' GROUP BY  TagClientGroup.Description, TFEntities.Code, TFEntities.Name

    ' ORDER BY Tots, Profit DESC
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableYTD') Is Not Null)
    ' Begin
    ' Drop Table #TempTableYTD
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWCurr10
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd5to9and11') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd5to9and11
    ' End
    ' If(OBJECT_ID('tempdb..#TempUninvoiced') Is Not Null)
    ' Begin
    ' Drop Table #TempUninvoiced
    ' End
    ' If(OBJECT_ID('tempdb..#TempIWytd10') Is Not Null)
    ' Begin
    ' Drop Table #TempIWytd10
    ' End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E17_ServiceFeeAnalysis(ByRef mReport As ReportsCollection) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "USE TravelForceCosmos
    'SELECT dt.Code
    '     , dt.Series
    '     , d.Number
    '     , TFEntities.Code
    '     , TFEntities.Name
    '     , d.IssueDate
    '     , Currencies.ISOAlphabetic AS CurrencyCode
    '     , di.ServiceFee + di.ServiceFeeVatAmount AS InvoiceServiceFee
    '     , sfasum.Amount AS CardServFeeAmount
    '     , di.ServiceFee + di.ServiceFeeVatAmount - sfasum.Amount AS SFDifference
    '     , ISNULL(sfaTFee.Description, '') AS TfeeD
    '      , ISNULL(sfaTFee.Amount, 0) AS TFee
    '      , ISNULL(sfaTFeeDom.Description, '') AS TfeeDomD
    '      , ISNULL(sfaTFeeDom.Amount, 0) AS TFeeDom
    '      , ISNULL(sfaIW5.Description, '') AS IW5D
    '      , ISNULL(sfaIW5.Amount, 0) AS IW5
    '      , ISNULL(sfaIW6.Description, '') AS IW6D
    '      , ISNULL(sfaIW6.Amount, 0) AS IW6
    '      , ISNULL(sfaIW7.Description, '') AS IW7D
    '      , ISNULL(sfaIW7.Amount, 0) AS IW7
    '      , ISNULL(sfaIW8.Description, '') AS IW8D
    '      , ISNULL(sfaIW8.Amount, 0) AS IW8
    '      , ISNULL(sfaIW9.Description, '') AS IW9D
    '      , ISNULL(sfaIW9.Amount, 0) AS IW9
    '      , ISNULL(sfaIW10.Description, '') AS IW10D
    '      , ISNULL(sfaIW10.Amount, 0) AS IW10
    '      , ISNULL(sfaIW11.Description, '') AS IW11D
    '      , ISNULL(sfaIW11.Amount, 0) AS IW11
    'FROM Documents AS d 
    'INNER JOIN DocumentItems AS di 
    '    ON d.Id = di.DocumentsID 
    'INNER JOIN DocTypes AS dt 
    '    ON d.DocTypesID = dt.Id 
    'INNER JOIN (SELECT sfa.CommercialTransactionValueID, SUM(sfa.Amount) AS Amount
    '            FROM Documents AS d 
    '            INNER JOIN DocumentItems AS di 
    '                ON d.Id = di.DocumentsID 
    '            INNER JOIN ServiceFeeAnalysis AS sfa 
    '                ON di.CommercialTransactionValueID = sfa.CommercialTransactionValueID                               
    '            WHERE (d.IssueDate BETWEEN @FromDate AND @ToDate) 
    '              AND (d.DocStatusID = 41) 
    '              AND (d.IsCancellationDocument = 0)
    '            GROUP BY sfa.CommercialTransactionValueID) AS sfasum 
    '    ON di.CommercialTransactionValueID = sfasum.CommercialTransactionValueID 
    '    AND di.ServiceFee + di.ServiceFeeVatAmount <> sfasum.Amount
    'LEFT JOIN Currencies
    'ON d.CurrencyID =Currencies.Id
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW5
    'ON di.commercialtransactionvalueid = sfaIW5.commercialtransactionvalueid AND sfaIW5.ServiceFeeTypeId =1
    'LEFT JOIN ServiceFeeAnalysis AS sfaTFee
    'ON di.commercialtransactionvalueid = sfaTFee.commercialtransactionvalueid AND sfaTFee.ServiceFeeTypeId =2
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW6
    'ON di.commercialtransactionvalueid = sfaIW6.commercialtransactionvalueid AND sfaIW6.ServiceFeeTypeId =3
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW7
    'ON di.commercialtransactionvalueid = sfaIW7.commercialtransactionvalueid AND sfaIW7.ServiceFeeTypeId =4
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW8
    'ON di.commercialtransactionvalueid = sfaIW8.commercialtransactionvalueid AND sfaIW8.ServiceFeeTypeId =5
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW9
    'ON di.commercialtransactionvalueid = sfaIW9.commercialtransactionvalueid AND sfaIW9.ServiceFeeTypeId =6
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW10
    'ON di.commercialtransactionvalueid = sfaIW10.commercialtransactionvalueid AND sfaIW10.ServiceFeeTypeId =7
    'LEFT JOIN ServiceFeeAnalysis AS sfaTFeeDom
    'ON di.commercialtransactionvalueid = sfaTFeeDom.commercialtransactionvalueid AND sfaTFeeDom.ServiceFeeTypeId =8
    'LEFT JOIN ServiceFeeAnalysis AS sfaIW11
    'ON di.commercialtransactionvalueid = sfaIW11.commercialtransactionvalueid AND sfaIW11.ServiceFeeTypeId =9
    'LEFT JOIN TFEntities
    'ON TFEntities.Id = d.CounterPartyID
    'WHERE (d.IsCancellationDocument = 0) 
    '  AND (d.DocStatusID = 41)
    'ORDER BY d.IssueDate

    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E18_AirTicketSales(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '	  , CASE WHEN ISNULL(Documents.IsCancellationDocument, 0) = 1 THEN ISNULL(DocTypes.CancelDocCode, '') ELSE ISNULL(DocTypes.Code, '') END AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    '      , ISNULL(Documents.DocStatusID, 0) AS DocStatusID
    '	  , ISNULL(CancDocType.Code + ' ' + CancDoc.Number, '') AS CancelsDoc
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    'LEFT JOIN TravelForceCosmos.dbo.Documents CancDoc WITH (NOLOCK)
    '	ON Documents.CancelsDocumentID = CancDoc.Id
    'LEFT JOIN TravelForceCosmos.dbo.DocTypes CancDocType WITH (NOLOCK)
    '    ON CancDocType.Id = CancDoc.DocTypesId
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND  CommercialTransactions.TransactionDate BETWEEN @FromDate AND @ToDate
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E19_DailyProfitReportInvoicesWithTicketNumber(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@WithTicket", SqlDbType.Bit).Value = mReport.BooleanOption1
    '        If mReport.ByClient = ReportsCollection.ClientReportType.AllClients Then
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '        ElseIf mReport.ByClient = ReportsCollection.ClientReportType.ByClient Then
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = 0
    '        Else
    '            sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = ""
    '            sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        End If

    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "USE TravelForceCosmos
    'If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    'Begin
    'Drop Table #TempTableCurr
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr5to9and11') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr5to9and11
    'End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    'If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    'Begin
    'Drop Table #TempTableClients
    'End
    '-- #TempIWCurr5to9and11 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr5to9and11
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis
    'WHERE ServiceFeeTypeID IN (1,3,4,5,6,9) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id 
    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399)
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID
    '-- #TempIWCurr10 --------------------------------------------------------------------------------------------------------------
    'SELECT  CommercialTransactionValueID, ServiceFeeTypeID, SUM(Amount) AS IWAmount
    'INTO #TempIWCurr10
    'FROM TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    'WHERE ServiceFeeTypeID IN (7) AND CommercialTransactionValueID IN (

    'SELECT DISTINCT CommercialTransactionValues.Id 

    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     RIGHT JOIN TravelForceCosmos.dbo.ServiceFeeAnalysis WITH (NOLOCK)
    '     ON CommercialTransactionValues.Id = ServiceFeeAnalysis.CommercialTransactionValueID
    ' ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id 

    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID)) 	
    '      AND TFEntities.Id IN (SELECT TFEntityId FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID IN(154,155))
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) 
    '      AND ServiceFeeAnalysis.Id IS NOT NULL
    '      AND CommercialTransactionValues.IsCost=0

    '      )
    'GROUP BY CommercialTransactionValueID, ServiceFeeTypeID

    '-- #TempTableCurr --------------------------------------------------------------------------------------------------------------
    '   SELECT TFEntities.Code AS Client
    '        , DocTypes.Code AS DocCode
    '        , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '        , Documents.Number as DocumentNumber
    '		, Documents.IssueDate
    '        , CASE WHEN @WithTicket = 1 THEN ISNULL(Airlines.IATACode, '') ELSE '' END AS Airline	 
    '        , CASE WHEN @WithTicket = 1 THEN CommercialTransactions.ProductNr ELSE '' END AS TicketNumber	 
    '        , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS FaceValueAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS TaxesAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS CommissionAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS DiscountAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS CancellationFeeAIR
    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS TFAIR


    '     , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END else 0 END)) AS NetPayableAIR
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '     CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN 
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END)) AS NetBuyAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN  ISNULL(CommercialTransactions.Pax, 0)  ELSE 0 END) AS PaxAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWAIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11AIR
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID = 1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10AIR


    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS FaceValueServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS TaxesServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS CommissionServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS DiscountServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS CancellationFeeServices
    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS TFServices

    '     , CONVERT(DECIMAL(18,2), SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ISNULL(ctvCost.Id,0) >=0 THEN
    '            (ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * Documents.CurrencyRate
    '            ELSE 0 END ELSE 0 END)) AS NetPayableServices
    '     , CONVERT(DECIMAL(18,2), -SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN CASE WHEN ctvCost.Id IS NOT NULL THEN
    '        (ISNULL(ctvCost.FaceValue, 0)
    '        + ISNULL(ctvCost.FVVatAmount, 0)
    '        + ISNULL(ctvCost.FaceValueExtra, 0)
    '        + ISNULL(ctvCost.FVXVatAmount, 0)
    '        + ISNULL(ctvCost.Taxes, 0)
    '        + ISNULL(ctvCost.TAXVatAmount, 0)
    '        + ISNULL(ctvCost.TaxesExtra, 0)
    '        + ISNULL(ctvCost.TAXXVatAmount, 0)
    '        + ISNULL(ctvCost.DiscountAmount, 0)
    '        + ISNULL(ctvCost.DISCVatAmount, 0)
    '        + ISNULL(ctvCost.CommissionAmount, 0)
    '        + ISNULL(ctvCost.COMVatAmount, 0)
    '        + ISNULL(ctvCost.ServiceFeeAmount, 0)
    '        + ISNULL(ctvCost.SFVatAmount, 0)
    '        + ISNULL(ctvCost.CancellationFeeAmount, 0)
    '        + ISNULL(ctvCost.CFVatAmount, 0))
    '        * ctvCost.Rate
    '        ELSE 0 END ELSE 0 END ))
    '        AS NetBuyServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN ISNULL(CommercialTransactions.Pax, 0) ELSE 0 END) AS PaxServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1
    '        THEN (ISNULL(IWTable.IWAmount,0) + ISNULL(IW10Table.IWAmount,0))* Documents.CurrencyRate ELSE 0 END) AS IWServices
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 1
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW05Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 3
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW06Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 4
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW07Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 5
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW08Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 6
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW09Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IWTable.ServiceFeeTypeID = 9
    '        THEN ISNULL(IWTable.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW11Services
    '      , SUM(CASE WHEN CommercialTransactions.ComTransactionTypeID <> 1 AND IW10Table.ServiceFeeTypeID = 7
    '        THEN ISNULL(IW10Table.IWAmount,0)* Documents.CurrencyRate ELSE 0 END) AS IW10Services
    '      , CommercialTransactionValues.Omit
    ' INTO #TempTableCurr
    ' FROM TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    ' INNER JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '     LEFT JOIN #TempIWCurr5to9and11 IWTable
    '     ON IWTable.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN #TempIWCurr10 IW10Table
    '     ON IW10Table.CommercialTransactionValueID = CommercialTransactionValues.Id
    '     LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '     ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'ON CommercialTransactions.Id = CommercialTransactionValues.CommercialTransactionID
    'LEFT JOIN [TravelForceCosmos].[dbo].AirTickets ATTick WITH (NOLOCK)
    'ON ATTick.DocumentNr = CommercialTransactions.CustomDescription1
    'LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    'ON Airlines.Id = ATTick.TicketingAirlineID
    ' LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues ctvCost WITH (NOLOCK)
    ' ON CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID
    '    AND ctvCost.IsCost=1
    ' INNER JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    ' ON CommercialTransactionValues.Id = DocumentItems.CommercialTransactionValueID
    ' RIGHT OUTER JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    INNER JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '        INNER JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '        ON Documents.CounterPartyID = TFEntities.Id
    '    ON DocTypes.Id = Documents.DocTypesID
    ' ON DocumentItems.DocumentsID = Documents.Id
    ' WHERE (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '      AND (@TagID = 0 OR TFEntities.ID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WITH (NOLOCK) WHERE TagID = @TagID)) 	
    '      AND (Documents.IsCancellationDocument = 0)
    '      AND (Documents.DocStatusID = 41)
    '      AND (Documents.DocTypesID NOT IN (74, 75))
    '      AND (Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)
    '      AND CommercialTransactionValues.Id IS NOT NULL
    '      AND DocTypes.DocCategoryID NOT IN (13,399) 

    ' GROUP BY TFEntities.Code, TFEntityDepartments.Name, DocTypes.Code, Documents.Number, Documents.IssueDate, ISNULL(Airlines.IATACode, ''), CommercialTransactions.ProductNr, CommercialTransactionValues.Omit
    '-- #TempTableClients --------------------------------------------------------------------------------------------------------------
    'SELECT DISTINCT TFEntities.Code
    'INTO #TempTableClients
    'FROM TravelForceCosmos.dbo.TFEntities
    'LEFT JOIN #TempTableCurr       ON #TempTableCurr.Client       = TFEntities.Code
    'WHERE (#TempTableCurr.Client IS NOT NULL)
    '  AND (#TempTableCurr.NetPayableAir<>0 OR #TempTableCurr.NetBuyAIR <> 0 OR #TempTableCurr.IWAIR <>0 OR #TempTableCurr.NetPayableAir <>0
    '    OR #TempTableCurr.NetPayableServices<>0 OR #TempTableCurr.NetBuyServices <> 0 OR #TempTableCurr.IWServices <>0 OR #TempTableCurr.NetPayableServices <>0)
    '-- Result Recordset --------------------------------------------------------------------------------------------------------------
    '  SELECT 
    '        TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , #TempTableCurr.Vessel
    '      , #TempTableCurr.IssueDate
    '      , #TempTableCurr.DocCode AS DocCode
    '      , #TempTableCurr.DocumentNumber AS DocumentNumber
    '      , #TempTableCurr.Airline AS Airline
    '      , ISNULL(#TempTableCurr.TicketNumber, '') AS TicketNumber
    ' -- AIR -------------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueAir),0) AS FaceValueAir
    '      , COALESCE(SUM(#TempTableCurr.TaxesAir),0) AS TaxesAir
    '      , COALESCE(SUM(#TempTableCurr.CommissionAir),0) AS CommissionAir
    '      , COALESCE(SUM(#TempTableCurr.DiscountAir),0) AS DiscountAir
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeAir),0) AS CancellationFeeAir
    '      , COALESCE(SUM(#TempTableCurr.TFAir),0) AS TFAir
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) AS NetPayableAir
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) AS NetBuyAIR
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) AS IW05AIR
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) AS IW06AIR
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) AS IW07AIR
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) AS IW08AIR
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) AS IW09AIR
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) AS IW11AIR
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) AS IW10AIR
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) AS IWAIR
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) AS ProfitAir
    '      , COALESCE(SUM(#TempTableCurr.PaxAIR),0) AS PaxAIR
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxAir)),0),0),0)) AS ProfitPerPaxAir
    '-- SERVICES --------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueServices),0) AS FaceValueServices
    '      , COALESCE(SUM(#TempTableCurr.TaxesServices),0) AS TaxesServices
    '      , COALESCE(SUM(#TempTableCurr.CommissionServices),0) AS CommissionServices
    '      , COALESCE(SUM(#TempTableCurr.DiscountServices),0) AS DiscountServices
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeServices),0) AS CancellationFeeServices
    '      , COALESCE(SUM(#TempTableCurr.TFServices),0) AS TFServices

    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayableServices
    '      , COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuyServices
    '      , COALESCE(SUM(#TempTableCurr.IW05Services),0) AS IW05Services
    '      , COALESCE(SUM(#TempTableCurr.IW06Services),0) AS IW06Services
    '      , COALESCE(SUM(#TempTableCurr.IW07Services),0) AS IW07Services
    '      , COALESCE(SUM(#TempTableCurr.IW08Services),0) AS IW08Services
    '      , COALESCE(SUM(#TempTableCurr.IW09Services),0) AS IW09Services
    '      , COALESCE(SUM(#TempTableCurr.IW11Services),0) AS IW11Services
    '      , COALESCE(SUM(#TempTableCurr.IW10Services),0) AS IW10Services
    '      , COALESCE(SUM(#TempTableCurr.IWServices),0) AS IWServices
    '      , COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS ProfitServices
    '      , COALESCE(SUM(#TempTableCurr.PaxServices),0) AS PaxServices
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxServices)),0),0),0)) AS ProfitPerPaxServices
    '-- TOTAL -----------
    '      , COALESCE(SUM(#TempTableCurr.FaceValueAir),0) + COALESCE(SUM(#TempTableCurr.FaceValueServices),0) AS FaceValue
    '      , COALESCE(SUM(#TempTableCurr.TaxesAir),0) + COALESCE(SUM(#TempTableCurr.TaxesServices),0) AS Taxes
    '      , COALESCE(SUM(#TempTableCurr.CommissionAir),0) + COALESCE(SUM(#TempTableCurr.CommissionServices),0) AS Commission
    '      , COALESCE(SUM(#TempTableCurr.DiscountAir),0) + COALESCE(SUM(#TempTableCurr.DiscountServices),0) AS Discount
    '      , COALESCE(SUM(#TempTableCurr.CancellationFeeAir),0) + COALESCE(SUM(#TempTableCurr.CancellationFeeServices),0) AS CancellationFee
    '      , COALESCE(SUM(#TempTableCurr.TFAir),0) + COALESCE(SUM(#TempTableCurr.TFServices),0) AS TF

    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) AS NetPayable
    '      , COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) + COALESCE(SUM(#TempTableCurr.NetBuyServices),0) AS NetBuy
    '      , COALESCE(SUM(#TempTableCurr.IW05AIR),0) + COALESCE(SUM(#TempTableCurr.IW05Services),0)  AS IW05
    '      , COALESCE(SUM(#TempTableCurr.IW06AIR),0) + COALESCE(SUM(#TempTableCurr.IW06Services),0)  AS IW06
    '      , COALESCE(SUM(#TempTableCurr.IW07AIR),0) + COALESCE(SUM(#TempTableCurr.IW07Services),0)  AS IW07
    '      , COALESCE(SUM(#TempTableCurr.IW08AIR),0) + COALESCE(SUM(#TempTableCurr.IW08Services),0)  AS IW08
    '      , COALESCE(SUM(#TempTableCurr.IW09AIR),0) + COALESCE(SUM(#TempTableCurr.IW09Services),0)  AS IW09
    '      , COALESCE(SUM(#TempTableCurr.IW11AIR),0) + COALESCE(SUM(#TempTableCurr.IW11Services),0)  AS IW11
    '      , COALESCE(SUM(#TempTableCurr.IW10AIR),0) + COALESCE(SUM(#TempTableCurr.IW10Services),0)  AS IW10
    '      , COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.IWServices),0)  AS IW
    '      , COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0) AS Profit
    '      , COALESCE(SUM(#TempTableCurr.PaxAir),0) + COALESCE(SUM(#TempTableCurr.PaxServices),0) AS Pax
    '      , CONVERT(DECIMAL(18,2),COALESCE((COALESCE(SUM(#TempTableCurr.NetPayableAir),0) - COALESCE(SUM(#TempTableCurr.NetBuyAIR),0) - COALESCE(SUM(#TempTableCurr.IWAIR),0) + COALESCE(SUM(#TempTableCurr.NetPayableServices),0) - COALESCE(SUM(#TempTableCurr.NetBuyServices),0) - COALESCE(SUM(#TempTableCurr.IWServices),0))/NULLIF(COALESCE(SUM(ABS(#TempTableCurr.PaxAir)),0) + COALESCE(SUM(ABS(#TempTableCurr.PaxServices)),0),0),0)) AS ProfitPerPax
    '      , #TempTableCurr.Omit AS Omit
    'FROM #TempTableClients
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    'ON #TempTableClients.Code = TFEntities.Code
    ' LEFT JOIN #TempTableCurr
    'ON #TempTableCurr.Client=TFEntities.Code
    ' GROUP BY TFEntities.Code, TFEntities.Name, #TempTableCurr.Vessel, #TempTableCurr.DocCode, #TempTableCurr.DocumentNumber, #TempTableCurr.IssueDate, Airline, TicketNumber, #TempTableCurr.Omit
    ' ORDER BY #TempTableCurr.IssueDate, #TempTableCurr.DocCode, #TempTableCurr.DocumentNumber, TicketNumber
    '-----------------------------------------------------------------------------------------------------------------------------------------
    ' If(OBJECT_ID('tempdb..#TempTableCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableCurr
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableClients') Is Not Null)
    ' Begin
    ' Drop Table #TempTableClients
    ' End
    ' If(OBJECT_ID('tempdb..#TempTableIWCurr') Is Not Null)
    ' Begin
    ' Drop Table #TempTableIWCurr
    ' End
    'If(OBJECT_ID('tempdb..#TempIWCurr10') Is Not Null)
    'Begin
    'Drop Table #TempIWCurr10
    'End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E20_HellasConfidence(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromIssueDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToIssueDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@IssueDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromInvoiceDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToInvoiceDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@InvoiceDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , CommercialTransactions.CustomDescription4 AS Details
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(Documents.IssueDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate ) AS NetPayable      
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType	  
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber

    'FROM TravelForceCosmos.dbo.CommercialTransactionValues
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents
    '    ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.DocTypes
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities
    '    ON TFEntities.Id = CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  WHERE CommercialTransactionValues.IsCost = 0
    '        AND CommercialTransactionValues.Omit <> 1
    '        AND ISNULL(AirTickets.Void, 0) <> 1
    '        AND Documents.Id IS NOT NULL
    '        AND  (@IssueDateChecked = 0 OR CommercialTransactions.TransactionDate BETWEEN @FromIssueDate AND @ToIssueDate)
    '        AND  (@InvoiceDateChecked = 0 OR Documents.IssueDate BETWEEN @FromInvoiceDate AND @ToInvoiceDate)
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    'ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E21_ReportByVerifiedUser(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromVerifyDate", SqlDbType.DateTime).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToVerifyDate", SqlDbType.DateTime).Value = DateAdd(DateInterval.Hour, 24, mReport.Date1To)
    '        sqlComm.Parameters.Add("@VerifiedUserName", SqlDbType.NVarChar, 254).Value = mReport.GroupList
    '        sqlComm.CommandType = CommandType.Text
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#Temp') Is Not Null)
    'Begin
    'Drop Table #Temp
    'End

    'SELECT TFEntities.Code AS ClientCode, TFEntities.Name AS ClientName, Tags.[Description] AS OpsGroup
    'INTO #Temp
    'FROM      [TravelForceCosmos].[dbo].TFEntities AS TFEntities
    'LEFT JOIN [TravelForceCosmos].[dbo].[TFEntityTags] AS TFEntityTags
    'LEFT JOIN [TravelForceCosmos].[dbo].Tags AS Tags
    'ON TagId = Tags.Id
    'ON TFEntities.Id = TFEntityTags.TFEntityId
    'WHERE TagGroupId = 149

    'USE AmadeusReports

    'SELECT ISNULL([doPCC], '') AS [PCC]
    '      ,ISNULL([doGDS], '') AS [GDS]
    '      ,ISNULL([doPNR], '') AS [PNR]
    '      ,ISNULL(PNRFinisherUsers.pfAgentName, doPCC + '-' + doUserGdsId) AS [Booked By]
    '      ,ISNULL([doDateLogged], '') AS [Date Logged]
    '      ,COALESCE(UsersVerified.pfAgentName, doVerifiedUserId, '') AS [Verified By]
    '      ,ISNULL([doVerifiedDate], '') AS [Date Verified]
    '      ,ISNULL([doVerificationReason], '') AS [Verification]
    '      ,ISNULL(DATEDIFF(MINUTE,doDateLogged, doVerifiedDate)/60.00, 0) AS HoursDiff
    '      ,ISNULL([doClientCode], '') AS [Client Code]
    '      ,ISNULL(#Temp.ClientName, '') AS [Client Name]
    '      ,ISNULL(#Temp.OpsGroup, '') AS [Client Group]
    '      ,[doPaxName] AS [Pax Name]
    '      ,[doItinerary] AS [Itinerary]
    '      ,[doTotal] AS [PNR Total]
    '      ,[doDownsellTotal] AS [Downsell Total]
    '      ,[doFareBasis] AS [PNR Fare Basis]
    '      ,[doDownsellFareBasis] AS [Downsell Fare Basis]
    '      ,[doGDSCommand] AS [GDS Pricing]
    '      ,[doIssuingCarrier] AS [Issuing Carrier]
    '  FROM [AmadeusReports].[dbo].[DownsellPNRLog]
    '  LEFT JOIN PNRFinisherUsers
    '  ON doPCC = pfPCC AND doUserGdsId = pfUser
    '  LEFT JOIN PNRFinisherUsers UsersVerified
    '  ON UsersVerified.pfPCC + '-' + UsersVerified.pfUser = doVerifiedUserId
    '  LEFT JOIN #Temp
    '  ON doClientCode = #Temp.ClientCode
    '  WHERE  ((doVerifiedUserId <> 'PRICER' AND  doVerificationReason <> 'NEW PRICER ENTRY')
    '            AND doVerifiedDate BETWEEN @FromVerifyDate AND @ToVerifyDate
    '            AND (@VerifiedUserName = '' OR ISNULL(UsersVerified.pfAgentName, doVerifiedUserId) = @VerifiedUserName))
    '  ORDER BY doVerifiedDate
    'If(OBJECT_ID('tempdb..#Temp') IS NOT NULL)
    'Begin
    'Drop Table #Temp
    'End
    '"
    '        Return sqlComm
    '    End Function
    '    Public Function E22_Euronav(ByRef mReport As ReportsCollection) As SqlCommand ' ByVal CustomerCode As String, ByVal FromDate As Date, ByVal ToDate As Date) As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND  Documents.IssueDate BETWEEN @FromDate AND @ToDate
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    '        AND (Documents.DocStatusID = 41)
    '        AND (Documents.DocTypesID NOT IN (74, 75))
    '	    AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '		AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E23_SeaChefs(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End
    'If(OBJECT_ID('tempdb..#TempBU') Is Not Null)
    'Begin
    'Drop Table #TempBU
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    '	  , TFEntities.Id AS EntityId
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes
    '    ON DocTypes.Id = Documents.DocTypesID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '	  AND (@TagID = 0 OR TFEntities.Code = @ClientCode OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    '      AND DocTypes.UsesFSD = 1
    'GROUP BY Documents.Id, TFEntities.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT DISTINCT #TempDate.EntityId, Tags.Description,
    'ISNULL(SeaChefsBusinessUnits.SupplierNumber, '') AS SupplierNumber,
    'ISNULL(SeaChefsBusinessUnits.SupplierSite, '') AS SupplierSite,
    'ISNULL(SeaChefsBusinessUnits.TaxRegNum, '') AS TaxRegNum
    'INTO #TempBU
    'FROM #TempDate
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityTags
    'ON TFEntityTags.TFEntityID=#TempDate.EntityId
    'LEFT JOIN TravelForceCosmos.dbo.Tags
    'ON TFEntityTags.TagID = Tags.Id
    'LEFT JOIN ATPIData.dbo.SeaChefsBusinessUnits
    'ON Tags.Description = SeaChefsBusinessUnits.BusinessUnit
    'WHERE Tags.TagGroupID=160

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL(#TempBU.Description, '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , ISNULL(#TempBU.SupplierNumber, '') AS [Supplier Number]
    '	  , ISNULL(#TempBU.SupplierSite, '') AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(#TempBU.TaxRegNum, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + CommercialTransactions.CustomDescription3 
    '	  + '_' + FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy')
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , CASE WHEN LEN(CPV05.Value)=26 THEN CPV05.Value + '-00000000-0000-0000-000000-000000' ELSE ISNULL(CPV05.Value, '') END AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	TFEntities.Code AS ClientCode
    '	  , TFEntities.Name AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , CommercialTransactions.CustomDescription2 AS PNR
    '	  , CommercialTransactions.CustomDescription1 AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , CommercialTransactionValues.ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ActionLT.Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , 'CY I STD EU SERVICE' AS TaxClassificationCode
    'FROM #TempDate WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.Documents
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN #TempBU WITH (NOLOCK)
    '    ON TFEntities.Id = #TempBU.EntityId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End
    'If(OBJECT_ID('tempdb..#TempBU') Is Not Null)
    'Begin
    'Drop Table #TempBU
    'End
    '"
    '        Return sqlComm

    '    End Function
    '    Public Function E29_SeaChefsDetailed(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '	  AND (@TagID = 0 OR TFEntities.Code = @ClientCode OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    'GROUP BY Documents.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL((SELECT Tags.Description FROM TravelForceCosmos.dbo.TFEntityTags LEFT JOIN Tags ON TFEntityTags.TagID = Tags.ID
    '	    WHERE TFEntityTags.TFEntityID = TFEntities.ID and Tags.TagGroupID = 160 and Tags.ID BETWEEN 700 AND 706), '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , '300524' AS [Supplier Number]
    '	  , '300524-2' AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(TFEntities.TaxRegistrationCode, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription3, '') 
    '	  + '_' + ISNULL(FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy'), '') 
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , ISNULL(CPV05.Value, '') AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	ISNULL(TFEntities.Code, '') AS ClientCode
    '	  , ISNULL(TFEntities.Name, '') AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , ISNULL(CommercialTransactions.CustomDescription2, '') AS PNR
    '	  , ISNULL(Airlines.IATACode, '') AS AirlineCode
    '	  , ISNULL(CommercialTransactions.CustomDescription1, '') AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , ISNULL(CommercialTransactionValues.ProductTypeLT, '') AS ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(ActionLT.Id, 0) AS Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , ISNULL(CommercialTransactionValues.Remarks, '') AS Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    'FROM #TempDate WITH (NOLOCK)
    'LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End"
    '        Return sqlComm

    '    End Function
    '    Public Function E24_ProfitPerAgentTotals(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = 149
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT ISNULL(dbo.Tags.Description, '') AS GroupName   
    '     , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent  
    '	 , ISNULL(Salespersons.Name, '') as SalesPerson
    '     , dbo.TFEntities.Code                      
    '	 , dbo.TFEntities.Name                      
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Sales                  
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) 
    '	 + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Cost            
    '	 ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)             
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END)) AS Profit                                             
    '	 , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax                 
    '	 ,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                           
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))              
    '	 /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax                              
    '	 FROM dbo.CommercialTransactions WITH (NOLOCK)              
    '	 INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID  
    '	 LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '	 LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1             
    '	 LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'              
    '	 INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID              
    '	 RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)                   
    '	 INNER JOIN dbo.Documents WITH (NOLOCK)                        
    '	 INNER JOIN dbo.TFEntities WITH (NOLOCK)                            
    '	 LEFT JOIN dbo.TFEntityTags                                  
    '	 LEFT JOIN dbo.Tags         
    '	 ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID                             
    '	 ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)                        
    '	 ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id                    
    '	 ON dbo.DocTypes.Id = dbo.Documents.DocTypesID               
    '	 ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id    
    '	 WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'                   
    '	 AND (dbo.Documents.IsCancellationDocument = 0)                
    '	 AND (dbo.Documents.DocStatusID = 41)               
    '	 AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)               
    '	 AND dbo.CommercialTransactionValues.Id IS NOT NULL                 
    '	 AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL   
    '	 GROUP BY ISNULL(dbo.Tags.Description, '') 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 , dbo.TFEntities.Code                 
    '	 , dbo.TFEntities.Name    
    '	 ORDER BY ISNULL(dbo.Tags.Description, '')                 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 ,dbo.TFEntities.Code  "
    '        Return sqlComm

    '    End Function
    '    Public Function E25_ProfitPerAgentTransactions(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagGroup", SqlDbType.Int).Value = 149
    '        sqlComm.Parameters.Add("@FromCurr", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToCurr", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT ISNULL(dbo.Tags.Description, '') AS GroupName   
    '     , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent  
    '	 , ISNULL(Salespersons.Name, '') as SalesPerson
    '     , dbo.TFEntities.Code                      
    '	 , dbo.TFEntities.Name    
    '	 , Documents.IssueDate
    ' 	 , ISNULL(DocTypes.Code, '') + ' ' + ISNULL(Documents.Series, '') + ' ' + ISNULL(Documents.Number, '') AS DocNumber
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Sales                  
    '	 , CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) 
    '	 + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))                                           
    '	 AS Cost            
    '	 ,CONVERT(DECIMAL(18,2),SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)             
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END)) AS Profit                                             
    '	 , SUM(ISNULL(dbo.CommercialTransactions.Pax, 0)) AS Pax                 
    '	 ,ISNULL(CONVERT(DECIMAL(18,2), (SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(dbo.CommercialTransactionValues.FaceValue, 0)              
    '	 + ISNULL(dbo.CommercialTransactionValues.FVVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.FaceValueExtra, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.FVXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.Taxes, 0)                 
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.TaxesExtra, 0)            
    '	 + ISNULL(dbo.CommercialTransactionValues.TAXXVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.DiscountAmount, 0)        
    '	 + ISNULL(dbo.CommercialTransactionValues.DISCVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CommissionAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.COMVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.ServiceFeeAmount, 0)      
    '	 + ISNULL(dbo.CommercialTransactionValues.SFVatAmount, 0)                                            
    '	 + ISNULL(dbo.CommercialTransactionValues.CancellationFeeAmount, 0) 
    '	 + ISNULL(dbo.CommercialTransactionValues.CFVatAmount, 0)) * dbo.CommercialTransactionValues.Rate                                                 
    '	 ELSE 0 END)              
    '	 + SUM(CASE WHEN ctvCost.Id IS NOT NULL THEN                                           
    '	 (ISNULL(ctvCost.FaceValue, 0)              
    '	 + ISNULL(ctvCost.FVVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.FaceValueExtra, 0)        
    '	 + ISNULL(ctvCost.FVXVatAmount, 0)                                           
    '	 + ISNULL(ctvCost.Taxes, 0)                 
    '	 + ISNULL(ctvCost.TAXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.TaxesExtra, 0)            
    '	 + ISNULL(ctvCost.TAXXVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.DiscountAmount, 0)        
    '	 + ISNULL(ctvCost.DISCVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CommissionAmount, 0)      
    '	 + ISNULL(ctvCost.COMVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.ServiceFeeAmount, 0)      
    '	 + ISNULL(ctvCost.SFVatAmount, 0)                                            
    '	 + ISNULL(ctvCost.CancellationFeeAmount, 0) + ISNULL(ctvCost.CFVatAmount, 0)) * ctvCost.Rate                                                 
    '	 ELSE 0 END))              
    '	 /(NULLIF(SUM(ISNULL(dbo.CommercialTransactions.Pax, 0) ), 0))), 0) AS ProfitPerPax                              
    '	 FROM dbo.CommercialTransactions WITH (NOLOCK)              
    '	 INNER JOIN dbo.CommercialTransactionValues ON dbo.CommercialTransactions.Id = dbo.CommercialTransactionValues.CommercialTransactionID  
    '	 LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '	 LEFT JOIN dbo.CommercialTransactionValues ctvCost ON dbo.CommercialTransactionValues.CommercialTransactionID = ctvCost.CommercialTransactionID AND ctvCost.IsCost=1             
    '	 LEFT JOIN dbo.ServiceFeeAnalysis ON dbo.CommercialTransactionValues.Id = dbo.ServiceFeeAnalysis.CommercialTransactionValueID AND dbo.ServiceFeeAnalysis.Description='IW5'              
    '	 INNER JOIN dbo.DocumentItems WITH (NOLOCK) ON dbo.CommercialTransactionValues.Id = dbo.DocumentItems.CommercialTransactionValueID              
    '	 RIGHT OUTER JOIN dbo.DocTypes WITH (NOLOCK)                   
    '	 INNER JOIN dbo.Documents WITH (NOLOCK)                        
    '	 INNER JOIN dbo.TFEntities WITH (NOLOCK)                            
    '	 LEFT JOIN dbo.TFEntityTags                                  
    '	 LEFT JOIN dbo.Tags         
    '	 ON dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID                             
    '	 ON dbo.TFEntities.Id = dbo.TFEntityTags.TFEntityID AND dbo.TFEntityTags.TagID IN (SELECT Id FROM dbo.Tags WHERE dbo.Tags.TagGroupID=@TagGroup AND dbo.Tags.Id=dbo.TFEntityTags.TagID)                        
    '	 ON dbo.Documents.CounterPartyID = dbo.TFEntities.Id                    
    '	 ON dbo.DocTypes.Id = dbo.Documents.DocTypesID               
    '	 ON dbo.DocumentItems.DocumentsID = dbo.Documents.Id    
    '	 WHERE SUBSTRING(dbo.TFEntities.Code,1,1) <= '0'                   
    '	 AND (dbo.Documents.IsCancellationDocument = 0)                
    '	 AND (dbo.Documents.DocStatusID = 41)               
    '	 AND (dbo.Documents.IssueDate BETWEEN  @FromCurr AND @ToCurr)               
    '	 AND dbo.CommercialTransactionValues.Id IS NOT NULL                 
    '	 AND dbo.DocTypes.DocCategoryID NOT IN (13,399) --AND DocTypes.AccGeneratorsID IS NOT NULL   
    '	 GROUP BY ISNULL(dbo.Tags.Description, '') 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 , dbo.TFEntities.Code                 
    '	 , dbo.TFEntities.Name
    '	 , Documents.IssueDate
    '	 , DocTypes.Code
    '	 , Documents.Series
    '	 , Documents.Number
    '	 ORDER BY ISNULL(dbo.Tags.Description, '')                 
    '	 , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '')
    '	 , ISNULL(Salespersons.Name, '')
    '	 ,dbo.TFEntities.Code  "
    '        Return sqlComm

    '    End Function
    '    Public Function E28_OptimizationSavings(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "SELECT [doPNR] AS RecordLocator
    '      , '' AS TicketNumber
    '      ,[doPCC] AS PseudoCity
    '	  ,ISNULL(PNRFinisherUsers.pfAgentName, DownsellPNRLog.doUserGdsId) AS Consultant
    '      ,[doClientCode] AS AccountCode
    '	  ,ISNULL(TFEntities.Name, '') AS AccountName
    '	  ,SUBSTRING(doItinerary, 1,5) AS DateOfTravel
    '      ,SUBSTRING(doItinerary, 7,1000) AS IataItinerary
    '	  ,SUBSTRING(doItinerary, 11,2) AS PlateCarrier
    '	  , CASE WHEN CHARINDEX(' X ', doPaxName )>0 THEN
    '	    SUBSTRING(doPaxName, CHARINDEX(' X ', doPaxName) + 3, 1000)
    '		ELSE
    '	    LEN(REPLACE(doPaxName, '1.',''))-LEN(REPLACE(REPLACE(doPaxName, '1.',''), '.','')) + 1
    '		END AS NoOfPax
    '      ,[doTotal] AS FareAmount
    '	  ,[doVerifiedSavingsAmount] AS DownsellAmountPerPax
    '	  ,CASE WHEN CHARINDEX(' X ', doPaxName )>0 THEN
    '	    SUBSTRING(doPaxName, CHARINDEX(' X ', doPaxName) + 3, 1000)
    '		ELSE
    '	    LEN(REPLACE(doPaxName, '1.',''))-LEN(REPLACE(REPLACE(doPaxName, '1.',''), '.','')) + 1
    '		END * doVerifiedSavingsAmount AS DownsellAmountPerPNR
    '      , CASE WHEN doGDS = '1A' THEN DownsellPNRLog.doUserGdsId ELSE '' END AS AmadeusAgent
    '      , CASE WHEN doGDS = '1G' THEN DownsellPNRLog.doUserGdsId ELSE '' END AS GalileoAgent
    '      ,[doVerifiedUserId] AS ActionedBy
    '      ,[doFareBasis] AS OriginalFareBasis
    '      ,[doDownsellFareBasis] AS NewFareBasis
    '      ,[doGDSCommand] AS PricingCommand

    '  FROM [AmadeusReports].[dbo].[DownsellPNRLog]
    '  LEFT JOIN PNRFinisherUsers
    '  ON DownsellPNRLog.doPCC = PNRFinisherUsers.pfPCC AND DownsellPNRLog.doUserGdsId = PNRFinisherUsers.pfUser
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities
    '  ON doClientCode = TFEntities.Code
    '  WHERE doVerifiedsavingsamount >0
    '  AND CONVERT(DATE, doVerifiedDate) BETWEEN @FromDate AND @ToDate"
    '        Return sqlComm
    '    End Function
    '    Public Function E30_AirTicketsFullDetails(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@TagID", SqlDbType.Int).Value = mReport.TagID
    '        sqlComm.Parameters.Add("@ClientCode", SqlDbType.NVarChar, 20).Value = mReport.SelectedCustomer
    '        sqlComm.Parameters.Add("@FromDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@DateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked

    '        sqlComm.Parameters.Add("@InvoicedStatus", SqlDbType.Int).Value = mReport.OptionTriplet ' 0 = All 1 = Uninvoiced 2 = Invoiced
    '        sqlComm.Parameters.Add("@AirlineCodes", SqlDbType.NVarChar, 254).Value = mReport.TextEntry.Replace(vbCrLf, "|")
    '        sqlComm.CommandTimeout = 200
    '        sqlComm.CommandText = "SELECT  
    '      CommercialTransactions.TransactionDate AS IssueDate
    '      , TFEntities.Code AS ClientCode
    '      , TFEntities.Name AS ClientName
    '      , CASE WHEN CommercialTransactionValues.Omit = 1 THEN 'OMIT' ELSE '' END AS Omit
    '      , CASE WHEN ISNULL(AirTickets.Void, 0) = 1 THEN 'VOID' ELSE '' END AS Void
    '      , CommercialTransactions.CustomDescription2 AS PNR
    '      , CommercialTransactions.CustomDescription1 AS TicketNumber
    '      , CommercialTransactions.CustomDescription3 AS Passenger
    '      , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '      , ISNULL(LookupTable.Name, '') AS ProductType
    '      , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(DocTypes.Code, '') AS InvCode
    '      , ISNULL(Documents.Series, '') AS InvSeries
    '      , ISNULL(Documents.InternalNumber, 0) AS InvNumber
    '      , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS InvoiceDate 
    '      , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '      , ISNULL(CPV01.Value, '') AS BookedBy
    '      , ISNULL(CPV02.Value, '') AS Office
    '      , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '      , ISNULL(CPV05.Value, '') AS CostCentre
    '      , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '      , ISNULL(CPV13.Value, '') AS OPT
    '      , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , ISNULL(CPV15.Value, '') AS [AccountCode]
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPayable
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS NetPurchase
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '            + ISNULL(CommercialTransactionValues.FVVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS FaceValue
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.Taxes, 0)
    '            + ISNULL(CommercialTransactionValues.TAXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Taxes
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '            + ISNULL(CommercialTransactionValues.DISCVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Discount
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '            + ISNULL(CommercialTransactionValues.COMVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS Commission
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '            + ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS CancellationFee
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '            + ISNULL(CommercialTransactionValues.FVXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS FVExtra
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '            + ISNULL(CommercialTransactionValues.TAXXVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS TaxExtra
    '      , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '            + ISNULL(CommercialTransactionValues.SFVatAmount, 0))
    '            * CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS ServiceFee

    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , CommercialTransactionValues.Remarks
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN 'OTHER'
    '             ELSE CASE WHEN CommercialTransactionCards.TypeID = 0 
    '                       THEN 'AIR' 
    '                       ELSE 'SERVICES' 
    '                       END
    '             END AS TransactionType
    '      , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Airlines.IATACode, '') AS TicketingAirline
    '      , ISNULL(CommercialTransactions.CustomDescription4, '') AS Routing
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '      , ISNULL(CommercialTransactionValues.Reference, '') AS Reference
    '      , ISNULL(CommercialTransactions.FromDate, '') AS DepartureDate
    '      , ISNULL(CommercialTransactions.ToDate, '') AS ArrivalDate
    '      , ISNULL(ViewConnectedDocuments.ConnectedCode,'') +' ' + ISNULL(ViewConnectedDocuments.ConnectedSeries,'') + ' ' + ISNULL(ViewConnectedDocuments.ConnectedNumber, '') AS ConnectedDocument
    '      , ISNULL(Passengers.Remarks, '') AS PaxRemarks
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK)
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK)
    '    ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK)
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Passengers
    '    ON CommercialTransactions.CustomDescription3 = Passengers.Name AND CommercialTransactions.Id = Passengers.CommercialTransactionID AND Passengers.IsLeader = 1
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK)
    '    ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK)
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK)
    '    ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '    ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.ViewConnectedDocuments WITH (NOLOCK)
    '    ON Documents.Id = ViewConnectedDocuments.ConnectedDocumentID
    '    LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK)
    '    ON DocTypes.Id = Documents.DocTypesId
    'LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK)
    '    ON TFEntities.Id = CommercialTransactionValues.CommercialEntityID
    'LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK)
    '    ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK)
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK)
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK)
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK)
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK)
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK)
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK)
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV15 WITH (NOLOCK)
    '    ON CPV15.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV15.CustomPropertyID = 15 AND CPV15.TFEntityId = TFEntities.Id
    'LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK)
    '    ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    'LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK)
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    'LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK)
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  WHERE TFEntities.Code IS NOT NULL
    '        AND CommercialTransactionValues.IsCost = 0
    '        AND (@InvoicedStatus = 0 OR (Documents.Id IS NULL AND @InvoicedStatus=1) OR (Documents.Id IS NOT NULL AND @InvoicedStatus=2))
    '        AND (@DateChecked=0 OR (CommercialTransactions.TransactionDate BETWEEN @FromDate AND @ToDate))
    '        AND (@DepDateChecked=0 OR (CommercialTransactions.FromDate BETWEEN @FromDepDate AND @ToDepDate))
    '        AND (@ClientCode = '' OR TFEntities.Code = @ClientCode)
    '        AND (@TagID = 0 OR CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID = @TagID))
    '        AND (@AirlineCodes = '' OR CHARINDEX(Airlines.IATACode, @AirlineCodes)>0)
    ' ORDER BY TFEntities.Code, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1"
    '        Return sqlComm

    '    End Function
    '    Public Function E31_SeaChefsStatementCheck(ByRef mReport As ReportsCollection) As SqlCommand
    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.Parameters.Add("@FromInvDate", SqlDbType.Date).Value = mReport.Date1From
    '        sqlComm.Parameters.Add("@ToInvDate", SqlDbType.Date).Value = mReport.Date1To
    '        sqlComm.Parameters.Add("@InvDateChecked", SqlDbType.Bit).Value = mReport.Date1Checked
    '        sqlComm.Parameters.Add("@FromDepDate", SqlDbType.Date).Value = mReport.Date2From
    '        sqlComm.Parameters.Add("@ToDepDate", SqlDbType.Date).Value = mReport.Date2To
    '        sqlComm.Parameters.Add("@DepDateChecked", SqlDbType.Bit).Value = mReport.Date2Checked
    '        sqlComm.CommandTimeout = 120
    '        sqlComm.CommandText = "If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End

    'SELECT  Documents.Id
    '	  ,MAX(CommercialTransactions.FromDate) AS LastFromDate
    'INTO #TempDate
    'FROM TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    'WHERE TFEntities.Code IS NOT NULL
    '      AND CommercialTransactionValues.IsCost = 0
    '      AND Documents.Id IS NOT NULL
    '      AND (@InvDateChecked = 0 OR CONVERT(DATE, COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '')) BETWEEN @FromInvDate AND @ToInvDate)
    '	  AND CommercialTransactionValues.CommercialEntityID IN (SELECT TFEntityID FROM TravelForceCosmos.dbo.TFEntityTags WHERE TagID IN (154,155))
    '	  AND ISNULL(CommercialTransactionValues.Omit, 0) = 0
    '	  AND ISNULL(Documents.IsCancellationDocument, 0) = 0
    '	  AND Documents.DocStatusID=41
    'GROUP BY Documents.Id
    'HAVING @DepDateChecked=0 OR CONVERT(DATE, MAX(CommercialTransactions.ToDate)) BETWEEN @FromDepDate AND @ToDepDate

    'SELECT  '' AS [Changed]
    '      , '' AS [Row Status]
    '	  , '' AS [Invoice Header Identifier]
    '	  , ISNULL((SELECT Tags.Description FROM TravelForceCosmos.dbo.TFEntityTags LEFT JOIN Tags ON TFEntityTags.TagID = Tags.ID
    '	    WHERE TFEntityTags.TFEntityID = TFEntities.ID and Tags.TagGroupID = 160 and Tags.ID BETWEEN 700 AND 706), '') AS [Business Unit]
    '	  , ISNULL(Documents.InternalNumber, 0) AS [Invoice Number]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Invoice Currency]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Invoice Amount]
    '	  , COALESCE(CommercialTransactionValues.InvoiceDate,CommercialTransactions.EntryDate, '') AS [Invoice Date]
    '	  , 'ATPI Greece Travel Marine S.A' AS [Supplier]
    '	  , '300524' AS [Supplier Number]
    '	  , '300524-2' AS [Supplier Site]
    '      , ISNULL(Currencies.ISOAlphabetic, '') AS [Payment Currency]
    '	  , CASE WHEN CommercialTransactions.ActionTypeID = 335 THEN 'Standard' ELSE 'Credit Memo' END AS [Invoice Type]
    '	  , ISNULL(TFEntities.TaxRegistrationCode, '') AS [First-Party Tax Registration Number]
    '	  , ISNULL(DocumentItems.Line, 0) AS [Line]
    '	  , 'Item' AS [Type]
    '	  , CONVERT(DECIMAL(18,2),(ISNULL(CommercialTransactionValues.FaceValue, 0)
    '			+ ISNULL(CommercialTransactionValues.FVVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.FaceValueExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.FVXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.Taxes, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.TaxesExtra, 0)
    '			+ ISNULL(CommercialTransactionValues.TAXXVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DiscountAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.DISCVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CommissionAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.COMVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.ServiceFeeAmount, 0) 
    '			+ ISNULL(CommercialTransactionValues.SFVatAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CancellationFeeAmount, 0)
    '			+ ISNULL(CommercialTransactionValues.CFVatAmount, 0))
    '			* CommercialTransactionValues.Rate * (1-ISNULL(AirTickets.Void, 0))) AS [Amount]
    '      , ISNULL(Airlines.IATACode, '') 
    '	  + '_' + ISNULL(CPV01.Value, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription4, '')
    '	  + '_' + ISNULL(CommercialTransactions.CustomDescription3, '') 
    '	  + '_' + ISNULL(FORMAT(CommercialTransactions.FromDate, 'dd/MM/yyyy'), '') 
    '	  + '_' + ISNULL(CPV14.Value, '') AS [Description]
    '	  , CASE WHEN CommercialTransactionCards.TypeID = 0
    '	         THEN CASE WHEN CPV04.Value = 'MEDICAL' THEN '3104' ELSE '3102' END
    '             ELSE '3107'
    '			 END AS [Expense Type]
    '	  , ISNULL(CPV05.Value, '') AS [Distribution Combination]
    '	  , ISNULL(TFEntityDepartments.Name, '') AS Vessel
    '	  ,	ISNULL(TFEntities.Code, '') AS ClientCode
    '	  , ISNULL(TFEntities.Name, '') AS ClientName
    '	  , ISNULL(DocTypes.Code, '') AS InvCode
    '	  , ISNULL(Documents.Series, '') AS InvSeries
    '	  , ISNULL(CommercialTransactions.CustomDescription2, '') AS PNR
    '	  , ISNULL(Airlines.IATACode, '')  AS AirlineCode
    '	  , ISNULL(CommercialTransactions.CustomDescription1, '') AS TicketNumber
    '	  , (SELECT COUNT(*) FROM TravelForceCosmos.dbo.Passengers  WITH (NOLOCK) 
    '	     WHERE Passengers.CommercialTransactionID = CommercialTransactions.Id) AS PaxCount
    '	  , ISNULL(LookupTable.Name, '') AS ProductType
    '	  , ISNULL(CommercialTransactionValues.ProductTypeLT, '') AS ProductTypeLT
    '	  , ISNULL(ActionLT.Name, '') AS ActionType
    '      , ISNULL(ActionLT.Id, 0) AS Id
    '	  , ISNULL(CPV01.Value, '') AS BookedBy
    '	  , ISNULL(CPV02.Value, '') AS Office
    '	  , ISNULL(CPV04.Value, '') AS ReasonForTravel
    '	  , ISNULL(CPV11.Value, '') AS RequisitionNumber
    '	  , ISNULL(CPV13.Value, '') AS OPT
    '	  , ISNULL(CPV14.Value, '') AS [TRID-MarineFare]
    '      , CASE WHEN CommercialTransactionCards.TypeID IS NULL 
    '             THEN '' 
    '             ELSE CASE WHEN CommercialTransactionValues.Verified = 1 
    '                       THEN '' 
    '                       ELSE 'NOT VERIFIED' 
    '                       END
    '             END AS Verified
    '      , ISNULL(CommercialTransactionValues.Remarks, '') AS Remarks
    '	  , ISNULL(CommercialTransactionCards.RegNr, 0) AS RegNr
    '      , ISNULL(Salespersons.Name, '') as SalesPerson
    '      , ISNULL(CommercialTransactions.IssuePCC, '') + '/' + ISNULL(CommercialTransactions.IssueSalesmanString, '') AS IssuingAgent
    '      , ISNULL(CommercialTransactions.CreatorPCC, '') + '/' + ISNULL(CommercialTransactions.CreatorSalesmanString, '') AS CreatorAgent
    '	  , ISNULL(Airlines.IATACode, '') AS AirlineCode
    'FROM #TempDate WITH (NOLOCK)
    'LEFT JOIN TravelForceCosmos.dbo.Documents WITH (NOLOCK)
    '	ON #TempDate.Id = Documents.Id
    '  LEFT JOIN TravelForceCosmos.dbo.DocumentItems WITH (NOLOCK) 
    '	ON Documents.Id = DocumentItems.DocumentsId
    'LEFT JOIN TravelForceCosmos.dbo.CommercialTransactionValues WITH (NOLOCK) 
    '	ON DocumentItems.CommercialTransactionValueID = CommercialTransactionValues.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CommercialTransactions WITH (NOLOCK) 
    '	ON CommercialTransactions.Id = CommercialTransactionId
    '  LEFT JOIN CommercialTransactionCards WITH (NOLOCK) 
    '    ON CommercialTransactions.CardID = CommercialTransactionCards.Id
    '  LEFT JOIN TravelForceCosmos.dbo.AirTickets WITH (NOLOCK) 
    '	ON AirTickets.DocumentNr = CommercialTransactions.CustomDescription1
    '  LEFT JOIN TravelForceCosmos.dbo.Airlines WITH (NOLOCK) 
    '    ON Airlines.Id = AirTickets.TicketingAirlineID
    '  LEFT JOIN TravelForceCosmos.dbo.DocTypes WITH (NOLOCK) 
    '	ON DocTypes.Id = Documents.DocTypesId
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntities WITH (NOLOCK) 
    '	ON TFEntities.Id = CommercialEntityID
    '  LEFT JOIN TravelForceCosmos.dbo.TFEntityDepartments WITH (NOLOCK) 
    '	ON CommercialTransactionValues.CommercialEntityDepartmentID = TFEntityDepartments.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV01 WITH (NOLOCK) 
    '    ON CPV01.BookFileId = CommercialTransactions.CustomDescription5 AND CPV01.CustomPropertyID = 1 AND CPV01.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV02 WITH (NOLOCK) 
    '    ON CPV02.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV02.CustomPropertyID = 2 AND CPV02.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV04 WITH (NOLOCK) 
    '    ON CPV04.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV04.CustomPropertyID = 4 AND CPV04.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV05 WITH (NOLOCK) 
    '    ON CPV05.CTID = CommercialTransactionValues.CommercialTransactionId AND CPV05.CustomPropertyID = 5 AND CPV05.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV11 WITH (NOLOCK) 
    '    ON CPV11.BookFileId = CommercialTransactions.CustomDescription5 AND CPV11.CustomPropertyID = 11 AND CPV11.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV13 WITH (NOLOCK) 
    '    ON CPV13.BookFileId = CommercialTransactions.CustomDescription5 AND CPV13.CustomPropertyID = 13 AND CPV13.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.CustomPropertyValues CPV14 WITH (NOLOCK) 
    '    ON CPV14.BookFileId = CommercialTransactions.CustomDescription5 AND CPV14.CustomPropertyID = 14 AND CPV14.TFEntityId = TFEntities.Id
    '  LEFT JOIN TravelForceCosmos.dbo.LookupTable WITH (NOLOCK) 
    '	ON CommercialTransactionValues.ProductTypeLT = LookupTable.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Lookuptable ActionLT WITH (NOLOCK) 
    '    ON CommercialTransactions.ActionTypeID = ActionLT.Id
    '  LEFT JOIN TravelForceCosmos.dbo.SalesPersons WITH (NOLOCK) 
    '    ON CommercialTransactions.SalesmanID = SalesPersons.Id
    '  LEFT JOIN TravelForceCosmos.dbo.Currencies WITH (NOLOCK) 
    '    ON CommercialTransactionValues.CurrencyID = Currencies.Id
    'ORDER BY TFEntities.Code, ISNULL(Documents.InternalNumber, 0), DocumentItems.Line, CommercialTransactions.CustomDescription2, CommercialTransactions.CustomDescription1

    'If(OBJECT_ID('tempdb..#TempDate') Is Not Null)
    'Begin
    'Drop Table #TempDate
    'End"
    '        Return sqlComm

    '    End Function
    '    Public Function ClientList() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT Code, Name, [Code] + '-' + [Name] AS DispName  " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " WHERE IsClient=1  AND CanHaveCT=1 " &
    '                           " UNION Select Code, Name, [Name] + '-' + [Code] AS DispName  " &
    '                           " FROM [TravelForceCosmos].[dbo].[TFEntities] " &
    '                           " WHERE IsClient=1  AND CanHaveCT=1 " &
    '                           " ORDER BY DispName"
    '        Return sqlComm

    '    End Function
    '    Public Function ClientGroupsAll() As SqlCommand

    '        ClientGroupsAll = New SqlCommand
    '        ClientGroupsAll = mCnn.CreateCommand
    '        ClientGroupsAll.CommandText = "SELECT [Id]
    '                            ,[Description]
    '                            FROM [TravelForceCosmos].[dbo].[Tags]
    '                            WHERE TagGroupID IN (146)
    '                            ORDER BY Description"
    '    End Function
    '    Public Function ClientGroupsSeaChefs() As SqlCommand

    '        ClientGroupsSeaChefs = New SqlCommand
    '        ClientGroupsSeaChefs = mCnn.CreateCommand
    '        ClientGroupsSeaChefs.CommandText = "SELECT [Id]
    '                            ,[Description]
    '                            FROM [TravelForceCosmos].[dbo].[Tags]
    '                            WHERE TagGroupID IN (162)
    '                            ORDER BY Description"
    '    End Function
    '    Public Function BSPMonths() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT LEFT(CONVERT(NCHAR(8), [Date], 112), 6) AS BSPDate " &
    '                           " FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                           " GROUP BY LEFT(CONVERT(NCHAR(8), [Date], 112), 6) " &
    '                           " ORDER BY BSPDate DESC"
    '        Return sqlComm

    '    End Function
    '    Public Function TransactionYears() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = "SELECT MIN(YEAR(TransactionDate)) AS MinYear, MAX(YEAR(TransactionDate)) AS MaxYear
    '                            FROM [TravelForceCosmos].[dbo].[CommercialTransactions]"
    '        Return sqlComm

    '    End Function

    '    Public Function BSPForthnights() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = " SELECT Date AS BSPDate " &
    '                            " FROM [TravelForceCosmos].[dbo].[BSPTicket] " &
    '                            " GROUP BY [Date] " &
    '                            " ORDER BY BSPDate DESC"
    '        Return sqlComm

    '    End Function
    '    Public Function AgentGroups() As SqlCommand

    '        Dim sqlComm As New SqlCommand
    '        sqlComm = mCnn.CreateCommand
    '        sqlComm.CommandText = "SELECT '' AS AgentName
    '                            UNION
    '                            SELECT DISTINCT pfAgentName AS AgentName
    '                            FROM [AmadeusReports].[dbo].[PNRFinisherUsers]
    '                            ORDER BY AgentName"
    '        Return sqlComm
    '    End Function
    '    Public Function Reader(cmm As SqlCommand) As SqlDataReader

    '        Reader = cmm.ExecuteReader

    '    End Function
End Class
