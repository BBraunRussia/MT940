using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public static class MyMonth
    {
        private static string[] DigitMonth = new string[12] { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };

        public static string MonthToDigit(string month)
        {
            int index = 0;

            for (int i = 0; i < 12; i++)
            {
                if (DigitMonth[i] == month)
                    index = i;
            }

            return GetFormatedIndex(index);
        }

        private static string GetFormatedIndex(int index)
        {
            index++;
            string stringIndex = index.ToString();

            return index < 10 ? "0" + stringIndex : stringIndex;
        }
    }
}
