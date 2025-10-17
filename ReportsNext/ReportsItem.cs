using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public class ReportsItem
    {
        public enum DateInitValue
        {
            FirstPrevMonthToEnd = 1,
            FirstCurrMonthToToday = 2,
            FirstJanToToday = 3,
            FromToPrevDayOrFriday = 4,
            FirstCurrMonthToYesterday = 5,
            From5DaysAfterEOM=6
        }
        public enum ClientCodeSelect
        {
            None = 0,
            ClientCodeAndGroup = 1,
            ClientCodeAndSeachefsGroup = 2,
            ClientCodeOnly = 3
        }

        public int Index { get; set; }
        public ReportsCollection.DBConnection DBConnection { get; set; } = ReportsCollection.DBConnection.TravelForce;
        public string GroupName { get; set; }
        public string ReportName { get; set; }
        public string Date1Text { get; set; } = "";
        public bool Date1Optional { get; set; } = false;
        public DateInitValue Date1Init { get; set; }
        public string Date2Text { get; set; } = "";
        public bool Date2Optional { get; set; } = false;
        public DateInitValue Date2Init { get; set; }
        public bool Date2OnlyFrom { get; set; } = false;
        public ClientCodeSelect ClientCode { get; set; } = 0;
        public string InitialClientCode { get; set; } = "";
        public bool DomInt { get; set; } = false;
        public bool BSPMonth { get; set; } = false;
        //public bool BSPFortnight { get; set; } = false;
        public string TextEntry { get; set; } = "";
        public bool TextEntryMultiLine { get; set; } = false;
        public bool ReportYearMonth { get; set; } = false;
        public string CheckBoxText { get; set; } = "";
        public string GroupListText { get; set; } = "";
        public ReportsCollection.GroupListType GroupList { get; set; } = ReportsCollection.GroupListType.Undefined;
        public bool OptionsTriplet { get; set; } = false;
        public string Options0Text { get; set; } = "";
        public string Options1Text { get; set; } = "";
        public string Options2Text { get; set; } = "";
        public bool Hidden { get; set; } = false;
        public bool OptimizationMonth { get; set; } = false;
    }
}
