using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public class E28_ActionReasonItem
    {
        public string Reason { get; }
        public int Postpone { get; internal set; }
        public int Reject { get; internal set; }
        public int PostponeDelayTotal { get; internal set; }
        public int RejectDelayTotal { get; internal set; }
        public E28_ActionReasonPerHourItem[] PerHourItems { get; internal set; }
        readonly int[] hourindex = { 15, 16, 17, 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 };
        readonly int[] daypartindex = { 2, 2, 2, 2, 2, 2, 3, 3, 3, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2 };
        readonly uint[] colori = {0x7863be7b,
0x7870c386,
0x787dc891,
0x788ace9d,
0x7898d3a8,
0x78a5d8b4,
0x78b2debf,
0x78c0e3cb,
0x78cde9d6,
0x78daeee2,
0x78e8f3ed,
0x78f5f9f9,
0x78fcf6f9,
0x78fce9ec,
0x78fcdddf,
0x78fbd0d2,
0x78fbc3c6,
0x78fbb6b9,
0x78faa9ac,
0x78fa9d9f,
0x78fa9092,
0x78f98385,
0x78f97678,
0x78f8696b};
        decimal MinDelay
        {
            get
            {
                decimal mindelay = decimal.MaxValue;
                for (int i = 0; i < PerHourItems.Count(); i++)
                {
                    if (PerHourItems[i].DelayAverage() < mindelay) mindelay = PerHourItems[i].DelayAverage();
                }
                return mindelay;
            }
        }
        decimal MaxDelay
        {
            get
            {
                decimal maxdelay = decimal.MinValue;
                for (int i = 0; i < PerHourItems.Count(); i++)
                {
                    if (PerHourItems[i].DelayAverage() > maxdelay) maxdelay = PerHourItems[i].DelayAverage();
                }
                return maxdelay;
            }
        }
        public int Ranking(int index)
        {
            if (MaxDelay == MinDelay)
            {
                return 0;
            }
            else
            {
                return (int)((PerHourItems[index].DelayAverage() - MinDelay) / (MaxDelay - MinDelay) * 24);
            }
            
        }
        public uint ValueColor(int index)
        {
            int temp = Ranking(index);
            if (temp > 23) return colori[colori.GetUpperBound(0)];
            return colori[temp];
        }
        public int Total
        {
            get
            {
                return Postpone + Reject;
            }
        }
        public decimal AverageDelay()
        {
            if (Total == 0) return 0;
            return (decimal)(PostponeDelayTotal + RejectDelayTotal) / (decimal)Total;
        }
        public E28_ActionReasonItem(string reason)
        {
            Reason = reason;
            Postpone = 0;
            Reject = 0;
            PostponeDelayTotal = 0;
            RejectDelayTotal = 0;
            SetPerHourItem();
        }
        void SetPerHourItem()
        {
            PerHourItems = new E28_ActionReasonPerHourItem[24];
            int HourCount = 8;
            for (int i = 0; i < PerHourItems.Length; i++)
            {
                HourCount++;
                if (HourCount > 23) HourCount = 0;
                PerHourItems[i] = new E28_ActionReasonPerHourItem
                {
                    Index = i,
                    Hour = HourCount,
                    DayPart = daypartindex[HourCount],
                    Count = 0,
                    DelayTotal = 0
                };
            }
        }
        public void AddPostpone(int hour, int delay)
        {
            Postpone++;
            PostponeDelayTotal += delay;
            AddHourItem(hour, 1, delay);
        }
        public void AddReject(int hour, int delay)
        {
            Reject++;
            RejectDelayTotal += delay;
            AddHourItem(hour, 1, delay);
        }
        public void AddPostpone(int hour, int count, int delay)
        {
            Postpone += count;
            PostponeDelayTotal += delay;
            AddHourItem(hour, count, delay);
        }
        public void AddItem(E28_ActionReasonItem item)
        {
            Postpone += item.Postpone;
            PostponeDelayTotal += item.PostponeDelayTotal;
            Reject += item.Reject;
            RejectDelayTotal+= item.RejectDelayTotal;
            for (int i = 0; i < item.PerHourItems.Length; i++)
            {
                AddHourItem(i, item.PerHourItems[hourindex[i]].Count, item.PerHourItems[hourindex[i]].DelayTotal);
            }
        }

        public void AddReject(int hour, int count, int delay)
        {
            Reject += count;
            RejectDelayTotal += delay;
            AddHourItem(hour, count, delay);
        }
        void AddHourItem(int hour, int count, int delay)
        {
            if (hour >= 0 && hour <= 23)
            {
                PerHourItems[hourindex[hour]].Count += count;
                PerHourItems[hourindex[hour]].DelayTotal += delay;
            }
        }
    }
}
