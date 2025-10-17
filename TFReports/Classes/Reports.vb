'Option Strict On
'Option Explicit On
'Namespace Reports
'Public Class ReportsItem
'    Private Structure ClassProps
'        Dim Index As Integer
'        Dim GroupName As String
'        Dim ReportName As String
'        Dim Date1Text As String
'        Dim Date2Text As String
'        Dim ClientCode As Boolean
'        Dim DomInt As Boolean
'        Dim BSPMonth As Boolean
'        Dim BSPFortnight As Boolean
'        Dim TextEntry As String
'        Dim TextEntryMultiLine As Boolean
'        Dim ReportYearMonth As Boolean
'        Dim CheckBoxText As String
'        Dim OperationsGroup As Boolean
'    End Structure

'    Private mudtProps As ClassProps

'    Public Sub New(ByVal pIndex As Integer, ByVal pGroupName As String, ByVal pReportName As String, ByVal pDate1Text As String _
'                 , ByVal pDate2Text As String, ByVal pClientCode As Boolean, ByVal pDomInt As Boolean, ByVal pBSPYearMonth As Boolean _
'                 , ByVal pBSPDate As Boolean, ByVal pTextEntry As String, ByVal pTextEntryMultiLine As Boolean, ByVal pReportYearMonth As Boolean, ByVal pCheckBoxText As String _
'                 , ByVal pOperationGroup As Boolean)
'        With mudtProps
'            .Index = pIndex
'            .GroupName = pGroupName
'            .ReportName = pReportName
'            .Date1Text = pDate1Text
'            .Date2Text = pDate2Text
'            .ClientCode = pClientCode
'            .DomInt = pDomInt
'            .BSPMonth = pBSPYearMonth
'            .BSPFortnight = pBSPDate
'            .TextEntry = pTextEntry
'            .TextEntryMultiLine = pTextEntryMultiLine
'            .ReportYearMonth = pReportYearMonth
'            .CheckBoxText = pCheckBoxText
'            .OperationsGroup = pOperationGroup
'        End With
'    End Sub

'    Public ReadOnly Property Index As Integer
'        Get
'            Index = mudtProps.Index
'        End Get
'    End Property
'    Public ReadOnly Property GroupName As String
'        Get
'            GroupName = mudtProps.GroupName
'        End Get
'    End Property
'    Public ReadOnly Property ReportName As String
'        Get
'            ReportName = mudtProps.ReportName
'        End Get
'    End Property
'    Public ReadOnly Property Date1Text As String
'        Get
'            Return mudtProps.Date1Text
'        End Get
'    End Property
'    Public ReadOnly Property Date2Text As String
'        Get
'            Return mudtProps.Date2Text
'        End Get
'    End Property
'    Public ReadOnly Property ClientCode As Boolean
'        Get
'            Return mudtProps.ClientCode
'        End Get
'    End Property
'    Public ReadOnly Property DomInt As Boolean
'        Get
'            Return mudtProps.DomInt
'        End Get
'    End Property
'    Public ReadOnly Property BSPMonth As Boolean
'        Get
'            Return mudtProps.BSPMonth
'        End Get
'    End Property
'    Public ReadOnly Property BSPFortnight As Boolean
'        Get
'            Return mudtProps.BSPFortnight
'        End Get
'    End Property
'    Public ReadOnly Property TextEntry As String
'        Get
'            Return mudtProps.TextEntry
'        End Get
'    End Property
'    Public ReadOnly Property TextEntryMultiLine As Boolean
'        Get
'            Return mudtProps.TextEntryMultiLine
'        End Get
'    End Property
'    Public ReadOnly Property ReportYearMonth As Boolean
'        Get
'            Return mudtProps.ReportYearMonth
'        End Get
'    End Property
'    Public ReadOnly Property CheckBoxText As String
'        Get
'            Return mudtProps.CheckBoxText
'        End Get
'    End Property
'    Public ReadOnly Property OperationsGroup As Boolean
'        Get
'            Return mudtProps.OperationsGroup
'        End Get
'    End Property
'End Class
'Public Class ReportsCollection
'    Inherits Collections.Generic.Dictionary(Of Integer, ReportsItem)

'    Private Structure ClassProps
'        Dim DateFrom As Date
'        Dim DateTo As Date
'        Dim PrevDateFrom As Date
'        Dim PrevDateTo As Date
'        Dim MonthDomestic As Boolean
'        Dim BSPMonthDate As String
'        Dim BSPFortDate As Date
'        Dim SelectedCustomer As String
'        Dim ReportYear As Integer
'        Dim ReportMonth As Integer
'        Dim DiffAsPercentage As Boolean
'        Dim TextEntryItems() As String
'        Dim TextEntry As String
'    End Structure
'    Private mudtProps As ClassProps

'    Public Sub New()
'        MyBase.Add(0, New ReportsItem(0, "Client Reports", "00 Euronav", "Invoice Date", "", True, False, False, False, "", False, False, "", False))
'        MyBase.Add(2, New ReportsItem(2, "BSP", "02 BSP Month Report by airline", "", "", False, True, True, False, "", False, False, "", False))
'        MyBase.Add(3, New ReportsItem(3, "BSP", "03 BSP Fortnight Report by ticket", "", "", False, False, False, True, "", False, False, "", False))
'        MyBase.Add(4, New ReportsItem(4, "BSP", "04 Ticket Information", "", "", False, False, False, False, "Ticket Numbers (one per line)", True, False, "", False))
'        MyBase.Add(5, New ReportsItem(5, "Sales Report", "05 Client Turnover", "Invoice Date", "", False, False, False, False, "", False, False, "", False))
'        MyBase.Add(6, New ReportsItem(6, "Profit Report", "06 Profit Per Client with Budget Comparison", "", "", False, False, False, False, "", False, True, "Difference as %", False))
'        MyBase.Add(7, New ReportsItem(7, "Profit Report", "07 Profit per OPS Group", "Invoice Date", "", False, False, False, False, "", False, False, "", False))
'        MyBase.Add(8, New ReportsItem(8, "Profit Report", "08 Profit per Client Group", "Invoice Date", "", False, False, False, False, "", False, False, "", False))
'        MyBase.Add(9, New ReportsItem(9, "Profit Report", "09 Profit per OPS Group with Extra", "Invoice Date", "", False, False, False, False, "", False, False, "", False))
'        MyBase.Add(10, New ReportsItem(10, "Profit Report", "10 Profit per Client Group with Extra", "Invoice Date", "", False, False, False, False, "", False, False, "", False))
'        MyBase.Add(11, New ReportsItem(11, "Profit Report", "11 Profit Per Client Group with PY", "", "", False, False, False, False, "", False, True, "", False))
'        MyBase.Add(12, New ReportsItem(12, "Profit Report", "12 Profit Per Client Group with Budget Comparison", "", "", False, False, False, False, "", False, True, "Difference as %", False))
'        MyBase.Add(13, New ReportsItem(13, "Sales Report", "13 Ticket Status", "Ticket Issue Date", "Departure Date", False, False, False, False, "", False, False, "Uninvoiced", False))
'        MyBase.Add(14, New ReportsItem(14, "Sales Report", "14 Tickets by Airline and Operator Group", "From Departure Date", "", False, False, False, False, "Airline", False, False, "", True))

'        mudtProps.DateFrom = DateAdd(DateInterval.Month, -1, DateSerial(Year(Today), Month(Today), 1))
'        mudtProps.DateTo = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, mudtProps.DateFrom))
'        mudtProps.MonthDomestic = False
'        mudtProps.BSPFortDate = Date.MinValue
'        mudtProps.ReportYear = Year(Today)
'        mudtProps.ReportMonth = Month(Today)
'        mudtProps.DiffAsPercentage = False

'    End Sub
'    Public Property ReportYear As Integer
'        Get
'            Return mudtProps.ReportYear
'        End Get
'        Set(value As Integer)
'            mudtProps.ReportYear = value
'        End Set
'    End Property
'    Public Property ReportMonth As Integer
'        Get
'            Return mudtProps.ReportMonth
'        End Get
'        Set(value As Integer)
'            mudtProps.ReportMonth = value
'        End Set
'    End Property
'    Public Property DiffAsPercentage As Boolean
'        Get
'            Return mudtProps.DiffAsPercentage
'        End Get
'        Set(value As Boolean)
'            mudtProps.DiffAsPercentage = value
'        End Set
'    End Property
'    Public ReadOnly Property E12_FromCurr As Date
'        Get
'            E12_FromCurr = DateSerial(ReportYear, ReportMonth, 1)
'        End Get
'    End Property
'    Public ReadOnly Property E12_ToCurr As Date
'        Get
'            E12_ToCurr = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, E12_FromCurr))
'        End Get
'    End Property
'    Public ReadOnly Property E12_FromYTD As Date
'        Get
'            E12_FromYTD = DateSerial(ReportYear, 1, 1)
'        End Get
'    End Property
'    Public ReadOnly Property E12_ToYTD As Date
'        Get
'            E12_ToYTD = E12_ToCurr
'        End Get
'    End Property
'    Public ReadOnly Property E12_FromPYCurr As Date
'        Get
'            E12_FromPYCurr = DateSerial(ReportYear - 1, ReportMonth, 1)
'        End Get
'    End Property
'    Public ReadOnly Property E12_ToPYCurr As Date
'        Get
'            E12_ToPYCurr = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, E12_FromPYCurr))
'        End Get
'    End Property
'    Public ReadOnly Property E12_FromPYTD As Date
'        Get
'            E12_FromPYTD = DateSerial(ReportYear - 1, 1, 1)
'        End Get
'    End Property
'    Public ReadOnly Property E12_ToPYTD As Date
'        Get
'            E12_ToPYTD = E12_ToPYCurr
'        End Get
'    End Property
'    Public Property DateFrom As Date
'        Get
'            Return mudtProps.DateFrom
'        End Get
'        Set(value As Date)
'            mudtProps.DateFrom = value
'        End Set
'    End Property
'    Public Property DateTo As Date
'        Get
'            Return mudtProps.DateTo
'        End Get
'        Set(value As Date)
'            mudtProps.DateTo = value
'        End Set
'    End Property
'    Public ReadOnly Property FromYTD As Date
'        Get
'            FromYTD = DateSerial(Year(DateFrom), 1, 1)
'        End Get
'    End Property
'    Public ReadOnly Property ToYTD As Date
'        Get
'            ToYTD = DateAdd(DateInterval.Day, -1, DateSerial(Year(DateTo), Month(DateTo), 1))
'        End Get
'    End Property
'    Public ReadOnly Property FromPYTD As Date
'        Get
'            FromPYTD = DateAdd(DateInterval.Year, -1, FromYTD)
'        End Get
'    End Property
'    Public ReadOnly Property ToPYTD As Date
'        Get
'            ToPYTD = DateAdd(DateInterval.Year, -1, ToYTD)
'        End Get
'    End Property

'    Public Property PrevDateFrom As Date
'        Get
'            Return mudtProps.PrevDateFrom
'        End Get
'        Set(value As Date)
'            mudtProps.PrevDateFrom = value
'        End Set
'    End Property

'    Public Property PrevDateTo As Date
'        Get
'            Return mudtProps.PrevDateTo
'        End Get
'        Set(value As Date)
'            mudtProps.PrevDateTo = value
'        End Set
'    End Property

'    Public Property MonthDomestic As Boolean
'        Get
'            Return mudtProps.MonthDomestic
'        End Get
'        Set(value As Boolean)
'            mudtProps.MonthDomestic = value
'        End Set
'    End Property

'    Public Property BSPMonthDate As String
'        Get
'            Return mudtProps.BSPMonthDate
'        End Get
'        Set(value As String)
'            mudtProps.BSPMonthDate = value
'        End Set
'    End Property

'    Public Property BSPFortDate As Date
'        Get
'            Return mudtProps.BSPFortDate
'        End Get
'        Set(value As Date)
'            mudtProps.BSPFortDate = value
'        End Set
'    End Property

'    Public Property SelectedCustomer As String
'        Get
'            Return mudtProps.SelectedCustomer
'        End Get
'        Set(value As String)
'            mudtProps.SelectedCustomer = value
'        End Set
'    End Property
'    Public Property TextEntry As String
'        Get
'            Return mudtProps.TextEntry
'        End Get
'        Set(value As String)
'            mudtProps.TextEntry = value
'            SplitTextToItems()
'        End Set
'    End Property
'    Private Sub SplitTextToItems()
'        With mudtProps
'            .TextEntryItems = .TextEntry.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
'            For i As Integer = 0 To .TextEntryItems.GetUpperBound(0)
'                .TextEntryItems(i) = .TextEntryItems(i).Trim
'                If .TextEntryItems(i).Length <> 10 Or Not IsNumeric(.TextEntryItems(i)) Then

'                    Dim i1 As Integer = TextEntryItems(i).IndexOf(".")
'                    If i1 < 10 Then
'                        .TextEntryItems(i) = .TextEntryItems(i).Substring(i1 + 1).Trim
'                    End If
'                    i1 = .TextEntryItems(i).IndexOf(" ")
'                    If i1 > 9 Then
'                        .TextEntryItems(i) = .TextEntryItems(i).Substring(0, i1).Trim
'                    End If
'                    If .TextEntryItems(i).Length = 13 Then
'                        .TextEntryItems(i) = TextEntryItems(i).Substring(3)
'                    End If
'                    If .TextEntryItems(i).Length <> 10 Then
'                        .TextEntryItems(i) = ""
'                    End If
'                End If
'            Next
'        End With
'    End Sub
'    Public ReadOnly Property TextEntryItemsCount As Integer
'        Get
'            If Not mudtProps.TextEntryItems Is Nothing AndAlso IsArray(mudtProps.TextEntryItems) Then
'                Return mudtProps.TextEntryItems.Count
'            Else
'                Return 0
'            End If
'        End Get
'    End Property
'    Public ReadOnly Property TextEntryItems(ByVal Index As Integer) As String
'        Get
'            If Index >= mudtProps.TextEntryItems.GetLowerBound(0) And Index <= mudtProps.TextEntryItems.GetUpperBound(0) Then
'                Return mudtProps.TextEntryItems(Index)
'            Else
'                Return ""
'            End If
'        End Get
'    End Property
'End Class
'End Namespace