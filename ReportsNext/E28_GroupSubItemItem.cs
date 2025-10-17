using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public class E28_GroupSubItemItem
    {
        public string Group { get; set; }
        public string Name { get; set; }
        public E28_AgentItemValues Actioned { get; set; }
        public E28_AgentItemValues Postponed { get; set; }
        public E28_AgentItemValues Rejected { get; set; }
        E28_AgentItemValues tot;
        public E28_AgentItemValues Total
        {
            get
            {
                if (tot == null)
                {
                    tot = new E28_AgentItemValues
                    {
                        Count = Actioned.Count + Postponed.Count + Rejected.Count,
                        ActualSavings = Actioned.ActualSavings + Postponed.ActualSavings + Rejected.ActualSavings,
                        PotentialSaving = Actioned.PotentialSaving + Postponed.PotentialSaving + Rejected.PotentialSaving,
                        MinutesDelayTotal = Actioned.MinutesDelayTotal + Postponed.MinutesDelayTotal + Rejected.MinutesDelayTotal
                    };
                }
                return tot;
            }
        }
        public E28_GroupSubItemItem(string group, string name)
        {
            Group = group;
            Name = name;
            Actioned = new E28_AgentItemValues();
            Postponed = new E28_AgentItemValues();
            Rejected = new E28_AgentItemValues();            
        }
    }
    public class E28_AgentItemValues
    {
        public int Count { get; set; }
        public decimal ActualSavings { get; set; }
        public decimal PotentialSaving { get; set; }
        public int MinutesDelayTotal { get; set; }
        public decimal MinutesDelayAverage
        {
            get
            {
                if (Count == 0) return 0;
                return (decimal)MinutesDelayTotal / (decimal)Count;
            }
        }
    }
}
