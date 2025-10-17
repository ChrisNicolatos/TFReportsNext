using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public class ReportsCollection : Dictionary<int, ReportsItem>
    {
        public enum DBConnection
        {
            Undefined = 0,
            TravelForce = 1,
            TravelForceSQL18 = 2
        }
        public enum ClientReportType
        {
            AllClients = 0,
            ByClient = 1,
            ByGroup = 2
        }
        public enum GroupListType
        {
            Undefined = 0,
            OperatorsGroup = 1,
            Agents = 2
        }
        ClientReportType _ByClient;
        string _SelectedCustomer = null;
        long _TagID = 0;
        string _CustomerGroup = null;
        string[] _TextEntryItems = null;
        string _TextEntry = null;
        public ReportsCollection()
        {
            base.Clear();
            // TODO replace this with loading from a database 
            base.Add(0, new ReportsItem
            {
                Index = 0,
                GroupName = "Client Reports",
                ReportName = "00 Euronav",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = true
            });
            base.Add(2, new ReportsItem
            {
                Index = 2,
                GroupName = "BSP",
                ReportName = "02 BSP Month Report by airline",
                DomInt = true,
                BSPMonth = true,
                Hidden = true
            });
            base.Add(3, new ReportsItem
            {
                Index = 3,
                GroupName = "BSP",
                ReportName = "03 BSP Fortnight Report by ticket",
                BSPFortnight = true,
                Hidden = true
            });
            base.Add(4, new ReportsItem
            {
                Index = 4,
                GroupName = "BSP",
                ReportName = "04 Ticket Information",
                TextEntry = "Ticket Numbers one per line)",
                TextEntryMultiLine = true,
                Hidden = true
            });
            base.Add(5, new ReportsItem
            {
                Index = 5,
                GroupName = "Sales Report",
                ReportName = "05 Client Turnover",
                Date1Text = "Invoice Date",
                Hidden = true
            });
            base.Add(6, new ReportsItem
            {
                Index = 6,
                GroupName = "Profit Report",
                ReportName = "06 Profit Per Client with Budget Comparison",
                ReportYearMonth = true,
                CheckBoxText = "Difference as %",
                Hidden = true
            });
            base.Add(7, new ReportsItem
            {
                Index = 7,
                GroupName = "Profit Report",
                ReportName = "07 Profit per OPS Group",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(8, new ReportsItem
            {
                Index = 8,
                GroupName = "Profit Report",
                ReportName = "08 Profit per Client Group",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(9, new ReportsItem
            {
                Index = 9,
                GroupName = "Profit Report",
                ReportName = "09 Profit per OPS Group with Extra",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(10, new ReportsItem
            {
                Index = 10,
                GroupName = "Profit Report",
                ReportName = "10 Profit per Client Group with Extra",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(11, new ReportsItem
            {
                Index = 11,
                GroupName = "Profit Report",
                ReportName = "11 Profit Per Client Group with PY",
                ReportYearMonth = true,
                Hidden = false
            });
            base.Add(12, new ReportsItem
            {
                Index = 12,
                GroupName = "Profit Report",
                ReportName = "12 Profit Per Client Group with Budget Comparison",
                TextEntry = "Ops groups (blank if all)",
                TextEntryMultiLine = true,
                ReportYearMonth = true,
                CheckBoxText = "Difference as %",
                Hidden = false
            });
            base.Add(13, new ReportsItem
            {
                Index = 13,
                GroupName = "Sales Report",
                ReportName = "13 Ticket Status",
                Date1Text = "Ticket Issue Date",
                Date2Text = "Departure Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                CheckBoxText = "Uninvoiced",
                Hidden = true
            });
            base.Add(15, new ReportsItem
            {
                Index = 15,
                GroupName = "Profit Report",
                ReportName = "15 Daily Profit Report",
                Date1Text = "Invoice Date",
                Date1Init = ReportsItem.DateInitValue.FirstCurrMonthToYesterday,
                TextEntry = "Top x Clients",
                CheckBoxText = "With Customer Group Details",
                Hidden = true
            });
            base.Add(17, new ReportsItem
            {
                Index = 17,
                GroupName = "Sales Report",
                ReportName = "17 Service Fee Analysis",
                Date1Text = "From Issue Date",
                Hidden = true
            });
            base.Add(18, new ReportsItem
            {
                Index = 18,
                GroupName = "Sales Report",
                ReportName = "18 Air Ticket Sales",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                CheckBoxText = "Ignore OMIT/VOID",
                Hidden = false
            });
            base.Add(19, new ReportsItem
            {
                Index = 19,
                GroupName = "Profit Report",
                ReportName = "19  Profit Report Invoices with IW",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                CheckBoxText = "With Tickets",
                Hidden = false
            });
            base.Add(20, new ReportsItem
            {
                Index = 20,
                GroupName = "Client Reports",
                ReportName = "20 Hellas Confidence",
                Date1Text = "Issue Date",
                Date2Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(21, new ReportsItem
            {
                Index = 21,
                DBConnection = ReportsCollection.DBConnection.TravelForceSQL18,
                GroupName = "Optimize Reports",
                ReportName = "21 Report by Verified User",
                Date1Text = "Verified Date",
                GroupListText = "Agents",
                GroupList = GroupListType.Agents,
                Hidden = false
            });
            base.Add(22, new ReportsItem
            {
                Index = 22,
                GroupName = "Client Reports",
                ReportName = "22 Euronav",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                CheckBoxText = "Amount in Invoice Currency",
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                Hidden = false
            });
            base.Add(23, new ReportsItem
            {
                Index = 23,
                GroupName = "Sales Report",
                ReportName = "23 Sea Chefs",
                Date1Text = "Invoice Date",
                Date2Text = "Departure Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndSeachefsGroup,
                Date1Optional = true,
                Date2Optional = true,
                CheckBoxText = "Separate sheet per Booked By",
                Hidden = false
            });
            base.Add(24, new ReportsItem
            {
                Index = 24,
                GroupName = "Profit Report",
                ReportName = "24 Profit per agent - Totals",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(25, new ReportsItem
            {
                Index = 25,
                GroupName = "Profit Report",
                ReportName = "25 Profit per agent - Transactions",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(26, new ReportsItem
            {
                Index = 26,
                GroupName = "Credit Limit",
                ReportName = "26 Active clients only",
                Hidden = false
            });
            base.Add(27, new ReportsItem
            {
                Index = 27,
                GroupName = "Credit Limit",
                ReportName = "27 All clients",
                Hidden = false
            });
            base.Add(28, new ReportsItem
            {
                Index = 28,
                DBConnection = ReportsCollection.DBConnection.TravelForceSQL18,
                GroupName = "Optimize Reports",
                ReportName = "28 Optimization Actions",
                Date1Text = "Verified Date",
                Date1Init = ReportsItem.DateInitValue.FirstCurrMonthToToday,
                Hidden = false
            });
            base.Add(29, new ReportsItem
            {
                Index = 29,
                GroupName = "Sales Report",
                ReportName = "29 Sea Chefs detailed",
                Date1Text = "Invoice Date",
                Date2Text = "Departure Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndSeachefsGroup,
                CheckBoxText = "Include extra data fields",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(30, new ReportsItem
            {
                Index = 30,
                GroupName = "Sales Report",
                ReportName = "30 Air Ticket Full Details",
                Date1Text = "Issue Date",
                Date2Text = "Departure Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(31, new ReportsItem
            {
                Index = 31,
                GroupName = "Sales Report",
                ReportName = "31 Sea Chefs Check",
                Date1Text = "Invoice Date",
                Date2Text = "Departure Date",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = true
            });
            base.Add(32, new ReportsItem
            {
                Index = 32,
                GroupName = "Accounting QC",
                ReportName = "32 Check for Selling Price",
                Date1Text = "Ticket Issue Date",
                Date1Init = ReportsItem.DateInitValue.FromToPrevDayOrFriday,
                Hidden = false
            });
            base.Add(33, new ReportsItem
            {
                Index = 33,
                GroupName = "Gaslog",
                ReportName = "33 - 012212 Crew Preparation for invoicing",
                Date1Text = "Ticket Issue Date",
                Hidden = false
            });
            base.Add(35, new ReportsItem
            {
                Index = 35,
                GroupName = "Gaslog",
                ReportName = "35 - 012217 Crew Preparation for invoicing",
                Date1Text = "Ticket Issue Date",
                Hidden = false
            });
            base.Add(61, new ReportsItem
            {
                Index = 61,
                GroupName = "Gaslog",
                ReportName = "61 - 012629 Crew",
                Date1Text = "Ticket Issue Date",
                Hidden = false
            });
            base.Add(62, new ReportsItem
            {
                Index = 62,
                GroupName = "Gaslog",
                ReportName = "62 - 012629 Corporate",
                Date1Text = "Ticket Issue Date",
                Hidden = false
            });
            base.Add(36, new ReportsItem
            {
                Index = 36,
                GroupName = "Sales Report",
                ReportName = "36 Sea Chefs - All Units",
                Date1Text = "Invoice Date",
                Date2Text = "Departure Date",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(37, new ReportsItem
            {
                Index = 37,
                GroupName = "Profit Report",
                ReportName = "37 Profit Report Invoices with IW for all clients see 19)",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                CheckBoxText = "With Tickets",
                Hidden = false
            });
            base.Add(38, new ReportsItem
            {
                Index = 38,
                GroupName = "Sales Report",
                ReportName = "38 Air Ticket Sales-Gaslog",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Hidden = false
            });
            base.Add(39, new ReportsItem
            {
                Index = 39,
                GroupName = "Accounting QC",
                ReportName = "39 GDS Import Errors",
                CheckBoxText = "Errors only",
                Date2Optional = true,
                Hidden = false
            });
            base.Add(40, new ReportsItem
            {
                Index = 40,
                GroupName = "Profit Report",
                ReportName = "40 Profit Per Agent with Budget Comparison",
                ReportYearMonth = true,
                CheckBoxText = "Difference as %",
                Hidden = true
            });
            base.Add(41, new ReportsItem
            {
                Index = 41,
                GroupName = "Profit Report",
                ReportName = "41 Daily Profit Report with RINVA Analysis",
                Date1Text = "Invoice Date",
                Date1Init = ReportsItem.DateInitValue.FirstCurrMonthToYesterday,
                TextEntry = "Top x Clients",
                CheckBoxText = "With Customer Group Details",
                Hidden = true
            });
            base.Add(42, new ReportsItem
            {
                Index = 42,
                GroupName = "Sales Report",
                ReportName = "42 Air Ticket with FC",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = true
            });
            base.Add(43, new ReportsItem
            {
                Index = 43,
                GroupName = "Profit Report",
                ReportName = "43 Daily Profit Report with Provisional Analysis",
                Date1Text = "Invoice Date",
                Date1Init = ReportsItem.DateInitValue.FirstCurrMonthToYesterday,
                Date2Text="Provisional Departure Date limit",
                Date2Init=ReportsItem.DateInitValue.From5DaysAfterEOM,
                Date2Optional = true,
                Date2OnlyFrom = true,
                TextEntry = "Top x Clients",
                CheckBoxText = "With Customer Group Details",
                Hidden = false
            });
            base.Add(45, new ReportsItem
            {
                Index = 45,
                GroupName = "Sales Report",
                ReportName = "45 Air Ticket Sales - All Fields",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Hidden = false
            });
            base.Add(46, new ReportsItem
            {
                Index = 46,
                GroupName = "Profit Report",
                ReportName = "19a Profit Report Transaction Totals",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(47, new ReportsItem
            {
                Index = 47,
                GroupName = "Profit Report",
                ReportName = "47 Daily Profit Report Totals Only",
                Date1Text = "Invoice Date",
                CheckBoxText = "Do not include Omit",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(48, new ReportsItem
            {
                Index = 48,
                GroupName = "Profit Report",
                ReportName = "48 Daily Profit Report Totals per Invoice",
                Date1Text = "Invoice Date",
                CheckBoxText = "Do not include Omit",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(49, new ReportsItem
            {
                Index = 49,
                DBConnection = ReportsCollection.DBConnection.TravelForceSQL18,
                GroupName = "Optimize Reports",
                ReportName = "49 Optimization Monthly Report",
                Date1Text = "Verified Date",
                Hidden = false
            });
            base.Add(50, new ReportsItem
            {
                Index = 50,
                DBConnection = ReportsCollection.DBConnection.TravelForceSQL18,
                GroupName = "Optimize Reports",
                ReportName = "50 Optimization Annual Report by Month",
                OptimizationMonth = true,
                Hidden = false
            });
            base.Add(51, new ReportsItem
            {
                Index = 51,
                GroupName = "Profit Report",
                ReportName = "51 Daily Profit Totals Per Category - PAVLOS",
                Date1Text = "Invoice Date",
                CheckBoxText = "Do not include Omit",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(52, new ReportsItem
            {
                Index = 52,
                GroupName = "Gaslog",
                ReportName = "52 Gaslog Monthly Statement",
                Date1Text = "Invoice Date",
                Hidden = false
            });
            base.Add(53, new ReportsItem
            {
                Index = 53,
                GroupName = "Sales Report",
                ReportName = "53 Sea Chefs - Invoices by Departure Date",
                Date1Text = "Invoice Date",
                Date2Text = "Departure Date",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(54, new ReportsItem
            {
                Index = 54,
                GroupName = "Sales Report",
                ReportName = "54 Client Statement",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                //OptionsTriplet = true,
                //Options0Text = "All",
                //Options1Text = "Uninvoiced Only",
                //Options2Text = "Invoiced Only",
                CheckBoxText = "Include Omit",
                Hidden = false
            });
            base.Add(55, new ReportsItem
            {
                Index = 55,
                GroupName = "Sales Report",
                ReportName = "55 Safety",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = true
            });
            base.Add(56, new ReportsItem
            {
                Index = 56,
                GroupName = "Standard Data",
                ReportName = "56 Clients per Client Group and Agent Group",
                Hidden = false
            });
            base.Add(57, new ReportsItem
            {
                Index = 57,
                GroupName = "Sales Report",
                ReportName = "57 TUI 030366",
                Date1Text = "Issue Date",
                Date1Init = ReportsItem.DateInitValue.FirstCurrMonthToToday,
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeOnly,
                InitialClientCode = "030366",
                CheckBoxText = "With Staff subtotals",
                Hidden = false
            });
            base.Add(58, new ReportsItem
            {
                Index = 58,
                GroupName = "Sales Report",
                ReportName = "58 Air Ticket Preferred Carriers",
                Date1Text = "Issue Date",
                Date2Text = "Departure Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Date1Optional = true,
                Date2Optional = true,
                Hidden = false
            });
            base.Add(59, new ReportsItem
            {
                Index = 59,
                GroupName = "Sales Report",
                ReportName = "59 Air Tickets by Invoice Date",
                Date1Text = "Invoice Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(60, new ReportsItem
            {
                Index = 60,
                GroupName = "Sales Report",
                ReportName = "60 Report for Lowest Class",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(63, new ReportsItem
            {
                Index = 63,
                GroupName = "Sales Report",
                ReportName = "63 Air Ticket Sales - Temenos",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Hidden = false
            });
            base.Add(64, new ReportsItem
            {
                Index = 64,
                GroupName = "Sales Report",
                ReportName = "64 Lowest Classes Report",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                Hidden = false
            });
            base.Add(65, new ReportsItem
            {
                Index = 65,
                GroupName = "Sales Report",
                ReportName = "65 Ops Sales",
                Date1Text = "Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                OptionsTriplet = true,
                Options0Text = "All",
                Options1Text = "Uninvoiced Only",
                Options2Text = "Invoiced Only",
                Hidden = false
            });
            base.Add(66, new ReportsItem
            {
                Index = 66,
                GroupName = "Suppliers Reports",
                ReportName = "66 Purchases per airline (TAMIS)",
                Date1Text = "Transaction Date",
                Hidden = false
            });
            base.Add(67, new ReportsItem
            {
                Index = 67,
                GroupName = "Sales Report",
                ReportName = "67 Columbia",
                Date1Text = "Invoice Issue Date",
                ClientCode = ReportsItem.ClientCodeSelect.ClientCodeAndGroup,
                TextEntry = "Airline codes one per line)",
                TextEntryMultiLine = true,
                //OptionsTriplet = true,
                //Options0Text = "All",
                //Options1Text = "Uninvoiced Only",
                //Options2Text = "Invoiced Only",
                //CheckBoxText = "Ignore OMIT/VOID",
                Hidden = false
            });
            ClearOptions();
        }
        public void ClearOptions()
        {
            Date1From = (new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)).AddMonths(-1);
            Date1To = Date1From.AddMonths(1).AddDays(-1);
            Date1Checked = true;
            Date2From = Date1From;
            Date2To = Date1To;
            Date2Checked = true;
            MonthDomestic = false;
            BSPFortDate = DateTime.MinValue;
            ReportYear = DateTime.Today.Year;
            ReportMonth = DateTime.Today.Month;
            BooleanOption1 = false;
            GroupList = "";
            _SelectedCustomer = "";
            _TagID = 0;
            _CustomerGroup = "";
            _TextEntryItems = null;
            _TextEntry = "";
        }
        public string GroupList { get; set; }
        public int ReportYear { get; set; }
        public int ReportMonth { get; set; }
        public bool BooleanOption1 { get; set; }
        public DateTime Date1From { get; set; }
        public DateTime Date1To { get; set; }
        public bool Date1Checked { get; set; }
        public DateTime Date2From { get; set; }
        public DateTime Date2To { get; set; }
        public bool Date2Checked { get; set; }
        public bool MonthDomestic { get; set; }
        public string BSPMonthDate { get; set; }
        public DateTime BSPFortDate { get; set; }
        public int OptionTriplet { get; set; }
        public ClientReportType ByClient
        {
            get
            {
                return _ByClient;
            }
        }

        public DateTime E12_FromCurr
        {
            get
            {
                return new DateTime(ReportYear, ReportMonth, 1);
            }
        }
        public DateTime E12_ToCurr
        {
            get
            {
                return E12_FromCurr.AddMonths(1).AddDays(-1);
            }
        }
        public DateTime E12_FromYTD
        {
            get
            {
                return new DateTime(ReportYear, 1, 1);
            }
        }
        public DateTime E12_ToYTD
        {
            get
            {
                return E12_ToCurr;
            }
        }
        public DateTime E12_FromPYCurr
        {
            get
            {
                return new DateTime(ReportYear - 1, ReportMonth, 1);
            }
        }
        public DateTime E12_ToPYCurr
        {
            get
            {
                return E12_FromPYCurr.AddMonths(1).AddDays(-1);
            }
        }
        public DateTime E12_FromPYTD
        {
            get
            {
                return new DateTime(ReportYear - 1, 1, 1);
            }
        }
        public DateTime E12_ToPYTD
        {
            get
            {
                return E12_ToPYCurr;
            }
        }

        public DateTime FromYTD
        {
            get
            {
                return new DateTime(Convert.ToDateTime(Date1From).Year, 1, 1);
            }
        }
        public DateTime ToYTD
        {
            get
            {
                return (new DateTime(Convert.ToDateTime(Date1To).Year, Convert.ToDateTime(Date1To).Month, 1)).AddDays(-1);
            }
        }
        public DateTime FromPYTD
        {
            get
            {
                return FromYTD.AddYears(-1);
            }
        }
        public DateTime ToPYTD
        {
            get
            {
                return ToYTD.AddYears(-1);
            }
        }
        public string SelectedCustomer
        {
            get
            {
                if (_ByClient != ClientReportType.ByClient)
                {
                    return "";
                }
                else
                {
                    return _SelectedCustomer;
                }
            }
            set
            {
                _SelectedCustomer = value;
                _ByClient = ClientReportType.ByClient;
            }
        }
        public long TagID
        {
            get
            {
                if (_ByClient != ClientReportType.ByGroup)
                {
                    return 0;
                }
                else
                {
                    return _TagID;
                }
            }
        }
        public string CustomerGroup
        {
            get
            {
                return _CustomerGroup;
            }
        }
        public void SetCustomerGroup(long pTagID, string GroupDescription)
        {
            _TagID = pTagID;
            _CustomerGroup = GroupDescription;
            _ByClient = ClientReportType.ByGroup;
        }
        public void SetAllClients()
        {
            _ByClient = ClientReportType.AllClients;
        }
        public string TextEntry
        {
            get
            {
                if (_TextEntry == null)
                {
                    return "";
                }
                else
                {
                    return _TextEntry;
                }
            }
            set
            {
                _TextEntry = value;
                SplitTextToItems();
            }
        }
        private void SplitTextToItems()
        {

            _TextEntryItems = _TextEntry.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i <= _TextEntryItems.GetUpperBound(0); i++)
            {
                _TextEntryItems[i] = _TextEntryItems[i].Trim();
                if (_TextEntryItems[i].Length != 10 || !_TextEntryItems[i].All(char.IsNumber))
                {

                    int i1 = TextEntryItems(i).IndexOf(".");
                    if (i1 < 10)
                    {
                        _TextEntryItems[i] = _TextEntryItems[i].Substring(i1 + 1).Trim();
                    }
                    i1 = _TextEntryItems[i].IndexOf(" ");
                    if (i1 > 9)
                    {
                        _TextEntryItems[i] = _TextEntryItems[i].Substring(0, i1).Trim();
                    }
                    if (_TextEntryItems[i].Length == 13)
                    {
                        _TextEntryItems[i] = TextEntryItems(i).Substring(3);
                    }
                    if (_TextEntryItems[i].Length != 10)
                    {
                        _TextEntryItems[i] = "";
                    }
                }
            }

        }
        public string TextEntryItems(int Index)
        {
            if (Index >= _TextEntryItems.GetLowerBound(0) && Index <= _TextEntryItems.GetUpperBound(0))
            {
                return _TextEntryItems[Index];
            }
            else
            {
                return "";
            }
        }
        public int TextEntryItemsCount
        {
            get
            {
                if (_TextEntryItems != null && _TextEntryItems.GetType().IsArray)
                {
                    return _TextEntryItems.Count();
                }
                else
                {
                    return 0;
                }
            }
        }

    }
}
