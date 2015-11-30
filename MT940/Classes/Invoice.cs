using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public class Invoice
    {
        private string _number;

        public string Number
        {
            get { return _number; }
            set { _number = value; }
        }

        public Invoice()
        {
            string day;

            if (DateTime.Now.DayOfYear / 10 == 0)
                day = "00" + DateTime.Now.DayOfYear.ToString();
            else if (DateTime.Now.DayOfYear / 100 == 0)
                day = "0" + DateTime.Now.DayOfYear.ToString();
            else
                day = DateTime.Now.DayOfYear.ToString();

            Number = day + "." + "001";
        }

        public string GetNumberFormated()
        {
            string number = Number;
            while (number[0] == '0')
                number = number.Remove(0, 1);

            number = number.Replace('.', '/');

            return number;
        }
    }
}
