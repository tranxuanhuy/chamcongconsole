using System;

namespace chamcong
{
    internal class DateTimeNote:IComparable<DateTimeNote>
    {
        private DateTime dateTime;
        private string v;

        public DateTimeNote()
        {
        }

        public DateTimeNote(DateTime dateTime, string v)
        {
            this.DateTime = dateTime;
            this.Note = v;
        }

        public DateTime DateTime { get => dateTime; set => dateTime = value; }
        public string Note { get => v; set => v = value; }

        public int CompareTo(DateTimeNote other)
        {
            return dateTime.CompareTo(other.dateTime);
        }
    }
}