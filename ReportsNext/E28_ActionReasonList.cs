using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public class E28_ActionReasonList : Dictionary<string, E28_ActionReasonItem>
    {
        public E28_ActionReasonItem Total { get; }
        public E28_ActionReasonList()
        {
            Total = new E28_ActionReasonItem("TOTAL");
        }
        public void AddPostponed(string reason, int hour, int delay)
        {
            if (!this.ContainsKey(reason))
            {
                E28_ActionReasonItem item = new E28_ActionReasonItem(reason);
                this.Add(reason, item);
            }
            this[reason].AddPostpone(hour, delay);
            Total.AddPostpone(hour, delay);
        }
        public void AddRejected(string reason, int hour, int delay)
        {
            if (!this.ContainsKey(reason))
            {
                E28_ActionReasonItem item = new E28_ActionReasonItem(reason);
                this.Add(reason, item);
            }
            this[reason].AddReject(hour, delay);
            Total.AddReject(hour, delay);
        }
        public SortedDictionary<string, E28_ActionReasonItem> GetSorted()
        {
            int count = 0;
            SortedDictionary<string, E28_ActionReasonItem> temp = new SortedDictionary<string, E28_ActionReasonItem>();
            foreach (E28_ActionReasonItem item in this.Values)
            {
                count++;
                string key = $"{(999999 - item.Total).ToString().PadLeft(20, '0')}:{count}";
                temp.Add(key, item);
            }
            return temp;
        }
    }
}
