using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports
{
    public class E28_GroupSubItemList : SortedDictionary<string, E28_GroupSubItemItem>
    {
        public E28_GroupSubItemItem Totals;
        public void AddActioned(string group, string name, decimal actualsavings, decimal potentialsavings, int delayminutes, string farecurrency, string downsellcurrency)
        {
            CheckExists(group, name);
            this[group].Actioned.Count++;
            this[group].Actioned.MinutesDelayTotal += delayminutes;
            this[$"{group}|{name}"].Actioned.Count++;
            this[$"{group}|{name}"].Actioned.MinutesDelayTotal += delayminutes;
            Totals.Actioned.Count++;
            Totals.Actioned.MinutesDelayTotal += delayminutes;
            if (farecurrency == downsellcurrency)
            {
                this[group].Actioned.ActualSavings += actualsavings;
                this[group].Actioned.PotentialSaving += potentialsavings;
                this[$"{group}|{name}"].Actioned.ActualSavings += actualsavings;
                this[$"{group}|{name}"].Actioned.PotentialSaving += potentialsavings;
                Totals.Actioned.ActualSavings += actualsavings;
                Totals.Actioned.PotentialSaving += potentialsavings;
            }

        }
        public void AddPostponed(string group, string name, decimal actualsavings, decimal potentialsavings, int delayminutes, string farecurrency, string downsellcurrency)
        {
            CheckExists(group, name);
            this[group].Postponed.Count++;
            this[group].Postponed.MinutesDelayTotal += delayminutes;
            this[$"{group}|{name}"].Postponed.Count++;
            this[$"{group}|{name}"].Postponed.MinutesDelayTotal += delayminutes;
            Totals.Postponed.Count++;
            Totals.Postponed.MinutesDelayTotal += delayminutes;
            if (farecurrency == downsellcurrency)
            {
                this[group].Postponed.ActualSavings += actualsavings;
                this[group].Postponed.PotentialSaving += potentialsavings;
                this[$"{group}|{name}"].Postponed.ActualSavings += actualsavings;
                this[$"{group}|{name}"].Postponed.PotentialSaving += potentialsavings;
                Totals.Postponed.ActualSavings += actualsavings;
                Totals.Postponed.PotentialSaving += potentialsavings;
            }

        }
        public void AddRejected(string group, string name, decimal actualsavings, decimal potentialsavings, int delayminutes, string farecurrency, string downsellcurrency)
        {
            CheckExists(group, name);
            this[group].Rejected.Count++;
            this[group].Rejected.MinutesDelayTotal += delayminutes;
            this[$"{group}|{name}"].Rejected.Count++;
            this[$"{group}|{name}"].Rejected.MinutesDelayTotal += delayminutes;
            Totals.Rejected.Count++;
            Totals.Rejected.MinutesDelayTotal += delayminutes; if (farecurrency == downsellcurrency)
            {
                this[group].Rejected.ActualSavings += actualsavings;
                this[group].Rejected.PotentialSaving += potentialsavings;
                this[$"{group}|{name}"].Rejected.ActualSavings += actualsavings;
                this[$"{group}|{name}"].Rejected.PotentialSaving += potentialsavings;
                Totals.Rejected.ActualSavings += actualsavings;
                Totals.Rejected.PotentialSaving += potentialsavings;
            }

        }
        void CheckExists(string group, string name)
        {
            if (!this.ContainsKey(group))
            {
                this.Add(group, new E28_GroupSubItemItem(group, ""));
            }
            if (!this.ContainsKey($"{group}|{name}"))
            {
                this.Add($"{group}|{name}", new E28_GroupSubItemItem(group, name));
            }
            if (Totals == null) Totals = new E28_GroupSubItemItem("", "");
        }
    }
}