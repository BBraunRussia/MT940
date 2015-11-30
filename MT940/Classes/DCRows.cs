using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public class DCRows
    {
        private List<DCRow> _list;

        public List<DCRow> List
        {
            get
            {
                Sort();
                return _list;
            }
        }

        public string Sum { get { return _list.Sum(item => item.Sum).ToString().Replace(",", ""); } }

        public DCRows()
        {
            _list = new List<DCRow>();
        }

        public void Add(DCRow row)
        {
            _list.Add(row);
        }

        private void Sort()
        {
            _list.Sort(delegate(DCRow d1, DCRow d2)
            { return d1.Sum.CompareTo(d2.Sum); });
        }
    }
}
