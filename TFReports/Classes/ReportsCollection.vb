Option Strict On
Option Explicit On
Public Class ReportsCollectionX
    'Inherits Collections.Generic.Dictionary(Of Integer, ReportsItemX)
    'Public Enum DBConnection
    '    Undefined = 0
    '    TravelForce = 1
    '    PNR = 2
    'End Enum
    'Public Enum ClientReportType
    '    AllClients = 0
    '    ByClient = 1
    '    ByGroup = 2
    'End Enum
    'Public Enum GroupListType
    '    Undefined = 0
    '    OperatorsGroup = 1
    '    Agents = 2
    'End Enum

    'Dim _ByClient As ClientReportType
    'Dim _SelectedCustomer As String
    'Dim _TagID As Integer
    'Dim _CustomerGroup As String
    'Dim _TextEntryItems() As String
    'Dim _TextEntry As String
    'Public Sub New()
    '    MyBase.Add(0, New ReportsItemX(0, ReportsCollectionX.DBConnection.TravelForce, "Client Reports", "00 Euronav", "Invoice Date", "", 1, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(2, New ReportsItemX(2, ReportsCollectionX.DBConnection.TravelForce, "BSP", "02 BSP Month Report by airline", "", "", 0, True, True, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(3, New ReportsItemX(3, ReportsCollectionX.DBConnection.TravelForce, "BSP", "03 BSP Fortnight Report by ticket", "", "", 0, False, False, True, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(4, New ReportsItemX(4, ReportsCollectionX.DBConnection.TravelForce, "BSP", "04 Ticket Information", "", "", 0, False, False, False, "Ticket Numbers (one per line)", True, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(5, New ReportsItemX(5, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "05 Client Turnover", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(6, New ReportsItemX(6, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "06 Profit Per Client with Budget Comparison", "", "", 0, False, False, False, "", False, True, "Difference as %", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(7, New ReportsItemX(7, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "07 Profit per OPS Group", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(8, New ReportsItemX(8, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "08 Profit per Client Group", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(9, New ReportsItemX(9, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "09 Profit per OPS Group with Extra", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(10, New ReportsItemX(10, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "10 Profit per Client Group with Extra", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(11, New ReportsItemX(11, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "11 Profit Per Client Group with PY", "", "", 0, False, False, False, "", False, True, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(12, New ReportsItemX(12, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "12 Profit Per Client Group with Budget Comparison", "", "", 0, False, False, False, "", False, True, "Difference as %", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(13, New ReportsItemX(13, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "13 Ticket Status", "Ticket Issue Date", "Departure Date", 0, False, False, False, "", False, False, "Uninvoiced", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(15, New ReportsItemX(15, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "15 Daily Profit Report", "Invoice Date", "", 0, False, False, False, "Top x Clients", False, False, "With Customer Group Details", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(17, New ReportsItemX(17, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "17 Service Fee Analysis", "From Issue Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(18, New ReportsItemX(18, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "18 Air Ticket Sales", "Issue Date", "", 1, False, False, False, "Airline codes (one per line)", True, False, "", "", GroupListType.Undefined, True, "All", "Uninvoiced Only", "Invoiced Only", False, False))
    '    MyBase.Add(19, New ReportsItemX(19, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "19 Profit Report Invoices with IW", "Invoice Date", "", 1, False, False, False, "", False, False, "With Tickets", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(20, New ReportsItemX(20, ReportsCollectionX.DBConnection.TravelForce, "Client Reports", "20 Hellas Confidence", "Issue Date", "Invoice Date", 1, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", True, True))
    '    MyBase.Add(21, New ReportsItemX(21, ReportsCollectionX.DBConnection.PNR, "Optimize Reports", "21 Report by Verified User", "Verified Date", "", 0, False, False, False, "", False, False, "", "Agents", GroupListType.Agents, False, "", "", "", False, False))
    '    MyBase.Add(22, New ReportsItemX(22, ReportsCollectionX.DBConnection.TravelForce, "Client Reports", "22 Euronav", "Invoice Date", "", 1, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(23, New ReportsItemX(23, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "23 Sea Chefs", "Invoice Date", "Departure Date", 2, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", True, True))
    '    MyBase.Add(24, New ReportsItemX(24, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "24 Profit per agent - Totals", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(25, New ReportsItemX(25, ReportsCollectionX.DBConnection.TravelForce, "Profit Report", "25 Profit per agent - Transactions", "Invoice Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(26, New ReportsItemX(26, ReportsCollectionX.DBConnection.TravelForce, "Credit Limit", "26 Active clients only", "", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(27, New ReportsItemX(27, ReportsCollectionX.DBConnection.TravelForce, "Credit Limit", "27 All clients", "", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(28, New ReportsItemX(28, ReportsCollectionX.DBConnection.PNR, "Optimize Reports", "28 Optimization Savings", "Verified Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, False))
    '    MyBase.Add(29, New ReportsItemX(29, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "29 Sea Chefs detailed", "Invoice Date", "Departure Date", 2, False, False, False, "", False, False, "Include extra data fields", "", GroupListType.Undefined, False, "", "", "", True, True))
    '    MyBase.Add(30, New ReportsItemX(30, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "30 Air Ticket Full Details", "Issue Date", "Departure Date", 1, False, False, False, "Airline codes (one per line)", True, False, "", "", GroupListType.Undefined, True, "All", "Uninvoiced Only", "Invoiced Only", True, True))
    '    MyBase.Add(31, New ReportsItemX(31, ReportsCollectionX.DBConnection.TravelForce, "Sales Report", "31 Sea Chefs Check", "Invoice Date", "Departure Date", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", True, True))
    '    MyBase.Add(32, New ReportsItemX(32, ReportsCollectionX.DBConnection.TravelForce, "Accounting QC", "32 Check for Selling Price", "Ticket Issue Date", "", 0, False, False, False, "", False, False, "", "", GroupListType.Undefined, False, "", "", "", False, True))

    '    Date1From = DateAdd(DateInterval.Month, -1, DateSerial(Year(Today), Month(Today), 1))
    '    Date1To = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, Date1From))
    '    Date1Checked = True
    '    Date2From = Date1From
    '    Date2To = Date1To
    '    Date2Checked = True
    '    MonthDomestic = False
    '    BSPFortDate = Date.MinValue
    '    ReportYear = Year(Today)
    '    ReportMonth = Month(Today)
    '    BooleanOption1 = False
    '    GroupList = ""
    'End Sub
    'Public Property GroupList As String
    'Public Property ReportYear As Integer
    'Public Property ReportMonth As Integer
    'Public Property BooleanOption1 As Boolean
    'Public Property Date1From As Date
    'Public Property Date1To As Date
    'Public Property Date1Checked As Boolean
    'Public Property Date2From As Date
    'Public Property Date2To As Date
    'Public Property Date2Checked As Boolean
    'Public Property MonthDomestic As Boolean
    'Public Property BSPMonthDate As String
    'Public Property BSPFortDate As Date
    'Public Property OptionTriplet As Integer
    'Public ReadOnly Property E12_FromCurr As Date
    '    Get
    '        E12_FromCurr = DateSerial(ReportYear, ReportMonth, 1)
    '    End Get
    'End Property
    'Public ReadOnly Property E12_ToCurr As Date
    '    Get
    '        E12_ToCurr = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, E12_FromCurr))
    '    End Get
    'End Property
    'Public ReadOnly Property E12_FromYTD As Date
    '    Get
    '        E12_FromYTD = DateSerial(ReportYear, 1, 1)
    '    End Get
    'End Property
    'Public ReadOnly Property E12_ToYTD As Date
    '    Get
    '        E12_ToYTD = E12_ToCurr
    '    End Get
    'End Property
    'Public ReadOnly Property E12_FromPYCurr As Date
    '    Get
    '        E12_FromPYCurr = DateSerial(ReportYear - 1, ReportMonth, 1)
    '    End Get
    'End Property
    'Public ReadOnly Property E12_ToPYCurr As Date
    '    Get
    '        E12_ToPYCurr = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, E12_FromPYCurr))
    '    End Get
    'End Property
    'Public ReadOnly Property E12_FromPYTD As Date
    '    Get
    '        E12_FromPYTD = DateSerial(ReportYear - 1, 1, 1)
    '    End Get
    'End Property
    'Public ReadOnly Property E12_ToPYTD As Date
    '    Get
    '        E12_ToPYTD = E12_ToPYCurr
    '    End Get
    'End Property

    'Public ReadOnly Property FromYTD As Date
    '    Get
    '        FromYTD = DateSerial(Year(Date1From), 1, 1)
    '    End Get
    'End Property
    'Public ReadOnly Property ToYTD As Date
    '    Get
    '        ToYTD = DateAdd(DateInterval.Day, -1, DateSerial(Year(Date1To), Month(Date1To), 1))
    '    End Get
    'End Property
    'Public ReadOnly Property FromPYTD As Date
    '    Get
    '        FromPYTD = DateAdd(DateInterval.Year, -1, FromYTD)
    '    End Get
    'End Property
    'Public ReadOnly Property ToPYTD As Date
    '    Get
    '        ToPYTD = DateAdd(DateInterval.Year, -1, ToYTD)
    '    End Get
    'End Property

    'Public ReadOnly Property ByClient As ClientReportType
    '    Get
    '        Return _ByClient
    '    End Get
    'End Property
    'Public Property SelectedCustomer As String
    '    Get
    '        If _ByClient <> ClientReportType.ByClient Then
    '            Return ""
    '        Else
    '            Return _SelectedCustomer
    '        End If
    '    End Get
    '    Set(value As String)
    '        _SelectedCustomer = value
    '        _ByClient = ClientReportType.ByClient
    '    End Set
    'End Property
    'Public ReadOnly Property TagID As Integer
    '    Get
    '        If _ByClient <> ClientReportType.ByGroup Then
    '            Return 0
    '        Else
    '            Return _TagID
    '        End If
    '    End Get
    'End Property
    'Public ReadOnly Property CustomerGroup As String
    '    Get
    '        Return _CustomerGroup
    '    End Get
    'End Property
    'Public Sub SetCustomerGroup(ByVal pTagID As Integer, ByVal GroupDescription As String)
    '    _TagID = pTagID
    '    _CustomerGroup = GroupDescription
    '    _ByClient = ClientReportType.ByGroup
    'End Sub
    'Public Sub SetAllClients()
    '    _ByClient = ClientReportType.AllClients
    'End Sub
    'Public Property TextEntry As String
    '    Get
    '        If _TextEntry Is Nothing Then
    '            Return ""
    '        Else
    '            Return _TextEntry
    '        End If
    '    End Get
    '    Set(value As String)
    '        _TextEntry = value
    '        SplitTextToItems()
    '    End Set
    'End Property

    'Private Sub SplitTextToItems()

    '    _TextEntryItems = _TextEntry.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    '    For i As Integer = 0 To _TextEntryItems.GetUpperBound(0)
    '        _TextEntryItems(i) = _TextEntryItems(i).Trim
    '        If _TextEntryItems(i).Length <> 10 Or Not IsNumeric(_TextEntryItems(i)) Then

    '            Dim i1 As Integer = TextEntryItems(i).IndexOf(".")
    '            If i1 < 10 Then
    '                _TextEntryItems(i) = _TextEntryItems(i).Substring(i1 + 1).Trim
    '            End If
    '            i1 = _TextEntryItems(i).IndexOf(" ")
    '            If i1 > 9 Then
    '                _TextEntryItems(i) = _TextEntryItems(i).Substring(0, i1).Trim
    '            End If
    '            If _TextEntryItems(i).Length = 13 Then
    '                _TextEntryItems(i) = TextEntryItems(i).Substring(3)
    '            End If
    '            If _TextEntryItems(i).Length <> 10 Then
    '                _TextEntryItems(i) = ""
    '            End If
    '        End If
    '    Next

    'End Sub
    'Public ReadOnly Property TextEntryItemsCount As Integer
    '    Get
    '        If _TextEntryItems IsNot Nothing AndAlso IsArray(_TextEntryItems) Then
    '            Return _TextEntryItems.Count
    '        Else
    '            Return 0
    '        End If
    '    End Get
    'End Property
    'Public ReadOnly Property TextEntryItems(ByVal Index As Integer) As String
    '    Get
    '        If Index >= _TextEntryItems.GetLowerBound(0) And Index <= _TextEntryItems.GetUpperBound(0) Then
    '            Return _TextEntryItems(Index)
    '        Else
    '            Return ""
    '        End If
    '    End Get
    'End Property
End Class
