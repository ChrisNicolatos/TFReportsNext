using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using System.Runtime.Remoting.Messaging;

namespace PNRHistory
{
    public class Remark
    {
        public string OrgId { get; set; }
        public string Id { get; set; }
        public string Action { get; set; }
        public string RType { get; set; }
        public string FreeFlow { get; set; }
    }
    public class Remarks : Dictionary<string, string>
    {
        public Dictionary<string, string> DummyRemarks { get; set; } = new Dictionary<string, string>();
        string RCRegex = @"^(?'OrgID'...)[\s\/](?'Id'...)\s(?'Action'AR|KO|XR)\/(?'RType'RC).+?\/\s*(?'FreeFlow'.+?)\r+$";
        string RXRegex = @"^(?'OrgID'...)[\s\/](?'Id'...)\s(?'Action'AR|KO|XR)\/(?'RType'RX)\s* (?'FreeFlow'.+?)\r+$";
        string RMRegex = @"^(?'OrgID'...)[\s\/](?'Id'...)\s(?'Action'AR|KO|XR)\/(?'RType'RM)\s* (?'FreeFlow'.+?)\r+$";
        string CLNRegex = @"^(?'OrgID'...)[\s\/](?'Id'...)\s(?'Action'AR|KO|XR)\/(?'RType'RM\* GRACE\/CLN/)(?'FreeFlow'.+?)\r+$";
        public Dictionary<string, string> ClientCode { get; set; } = new Dictionary<string, string>();
        public void AddRemarks(string pnr, string history)
        {
            string temp = "";
            temp = MatchRegexSale(RCRegex, history, temp);
            temp = MatchRegexSale(RXRegex, history, temp);
            if (temp != "" && !this.ContainsKey(pnr))
            {
                this.Add(pnr, temp);
            }

            temp = "";
            temp = MatchRegexDummy(RCRegex, history, temp);
            temp = MatchRegexDummy(RXRegex, history, temp);
            temp = MatchRegexDummy(RMRegex, history, temp);
            if (temp != "" && !this.ContainsKey(pnr))
            {
                DummyRemarks.Add(pnr, temp);
            }

            MatchClient(pnr, history);
        }
        void MatchClient(string pnr, string history)
        {
            Regex RegExMatch = new Regex(CLNRegex, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            MatchCollection RegExResult = RegExMatch.Matches(history);
            string temp = "";
            foreach (Match RCMatch in RegExResult)
            {
                if (RCMatch.Success)
                {
                    temp = $"{RCMatch.Groups["FreeFlow"].Value}";
                }
            }
            if (temp != "" && !ClientCode.ContainsKey(pnr))
            {
                ClientCode.Add(pnr, temp);
            }
        }
        string MatchRegexSale(string regex, string history, string allremarks)
        {
            Regex RegExMatch = new Regex(regex, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            MatchCollection RegExResult = RegExMatch.Matches(history);
            string temp = allremarks;
            foreach (Match RCMatch in RegExResult)
            {
                if (RCMatch.Success)
                {
                    if (RCMatch.Groups["Action"].Value != "XR")
                    {
                        if (temp.IndexOf(RCMatch.Groups["FreeFlow"].Value) < 0)
                        {
                            if (RCMatch.Groups["FreeFlow"].Value.IndexOf("SALE") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("SELL") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("EUR") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("USD") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("**") >= 0)
                            {
                                if (new Regex(@"\D*(?'Value'\d{2,})\D*").Match(RCMatch.Groups["FreeFlow"].Value).Success)
                                {
                                    if (temp.Length > 0) { temp += "|"; }
                                    temp += $"{RCMatch.Groups["FreeFlow"].Value}";
                                }
                            }
                        }

                    }
                    else
                    {
                        temp = temp.Replace(RCMatch.Groups["FreeFlow"].Value, "");
                        while (temp.StartsWith("|"))
                        {
                            temp = temp.Substring(1);
                        }
                        while (temp.EndsWith("|"))
                        {
                            temp = temp.Substring(0, temp.Length - 1);
                        }
                    }
                }

            }
            return temp;
        }
        string MatchRegexDummy(string regex, string history, string allremarks)
        {
            Regex RegExMatch = new Regex(regex, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            MatchCollection RegExResult = RegExMatch.Matches(history);
            string temp = allremarks;
            foreach (Match RCMatch in RegExResult)
            {
                if (RCMatch.Success)
                {
                    if (RCMatch.Groups["Action"].Value != "XR")
                    {
                        if (temp.IndexOf(RCMatch.Groups["FreeFlow"].Value) < 0)
                        {
                            if (RCMatch.Groups["FreeFlow"].Value.IndexOf("DUMMY") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("VISA") >= 0 || RCMatch.Groups["FreeFlow"].Value.IndexOf("PROTECT") >= 0)
                            {
                                if (RCMatch.Groups["FreeFlow"].Value.IndexOf("PLS INSRT IN SSR") < 0)
                                {
                                    if (temp.Length > 0) { temp += "|"; }
                                    temp += $"{RCMatch.Groups["FreeFlow"].Value}";
                                }
                            }
                        }

                    }
                    else
                    {
                        temp = temp.Replace(RCMatch.Groups["FreeFlow"].Value, "");
                        while (temp.StartsWith("|"))
                        {
                            temp = temp.Substring(1);
                        }
                        while (temp.EndsWith("|"))
                        {
                            temp = temp.Substring(0, temp.Length - 1);
                        }
                    }
                }

            }
            return temp;
        }
    }
}
