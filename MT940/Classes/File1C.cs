using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public class File1C
    {
        private ExcelDoc _excelBook;
        private string _currentCell;

        private string _date;
        private string _day;
        private string _month;
        private string _year;
        private string _compNumber;
        private string _incomeTail;
        private string _outcomeTail;
        private string _debetTotal;
        private string _creditTotal;

        private DCRows debet;
        private DCRows credit;

        public DCRows Debet { get { return debet; } }
        public DCRows Credit { get { return credit; } }

        public string CompNumber { get { return _compNumber; } }

        public string Day { get { return _day; } }
        public string MonthDigit { get { return MyMonth.MonthToDigit(_month); } }
        public string Year { get { return _year; } }
        
        public string IncomeTail { get { return _incomeTail.Replace('.', ','); } }
        public string OutcomeTail { get { return _outcomeTail.Replace('.', ','); } }

        public string DateFormated { get { return _year.Substring(2, 2) + MonthDigit + _day; } }

        public string CurrentCell
        {
            get { return _currentCell; }
            set { _currentCell = value; }
        }

        public File1C(ExcelDoc excelBook)
        {
            _excelBook = excelBook;

            _date = "";
            _day = "";
            _month = "";
            _year = "";
            _compNumber = "";
            _incomeTail = "";

            debet = new DCRows();
            credit = new DCRows();
        }
        
        public void Read()
        {
            ReadHeader();

            int i = GetIndexOnFirst();

            while (_excelBook.getValue("A" + i, "A" + i) != null)
            {
                DCRow dcRow = new DCRow();

                CurrentCell = "F" + i;
                dcRow.SetNumber(_excelBook.getValue("F" + i, "F" + i));

                CurrentCell = "D" + i;
                if (_excelBook.getValue("D" + i, "D" + i) != null)
                {
                    CurrentCell = "B" + (i + 1);
                    dcRow.SetOrdpWithoutDigit(_excelBook.getValue("B" + (i + 1), "B" + (i + 1)));

                    CurrentCell = "C" + (i + 2);
                    dcRow.SetBenm(_excelBook.getValue("C" + (i + 2), "C" + (i + 2)));

                    CurrentCell = "I" + i;
                    dcRow.SetCom(_excelBook.getValue("I" + i, "I" + i));

                    CurrentCell = "D" + i;
                    dcRow.SetSum(_excelBook.getValue("D" + i, "D" + i));

                    debet.Add(dcRow);
                }
                else
                {
                    CurrentCell = "B" + (i + 2);
                    dcRow.SetOrdp(_excelBook.getValue("B" + (i + 2), "B" + (i + 2)));

                    CurrentCell = "C" + (i + 1);
                    dcRow.SetBenmWithoutDigit(_excelBook.getValue("C" + (i + 1), "C" + (i + 1)));

                    CurrentCell = "I" + i;
                    dcRow.SetCom(_excelBook.getValue("I" + i, "I" + i));

                    CurrentCell = "E" + i;
                    dcRow.SetSum(_excelBook.getValue("E" + i, "E" + i));

                    credit.Add(dcRow);
                }

                i += 3;
            }

            ReadTails(i);
        }

        private void ReadHeader()
        {
            int i = 1;
            int indexAccountNumber = 1;

            while (i < 100)
            {
                if (_excelBook.getValue("A" + i, "A" + i) != null)
                {
                    if (_excelBook.getValue("A" + i, "A" + i).ToString().Replace("\n", "") == "по")
                    {
                        break;
                    }
                    if ((_excelBook.getValue("A" + i, "A" + i).ToString().Replace("\n", "") == "ВЫПИСКА ОПЕРАЦИЙ ПО ЛИЦЕВОМУ СЧЕТУ")
                        || (_excelBook.getValue("A" + i, "A" + i).ToString().Replace("\n", "") == "Счет:"))
                    {
                        indexAccountNumber = i;
                    }
                }
                i++;
            }

            if (_excelBook.getValue("B" + indexAccountNumber, "B" + indexAccountNumber) != null)
                _compNumber = _excelBook.getValue("B" + indexAccountNumber, "B" + indexAccountNumber).ToString().Replace("\n", "");
            if (_excelBook.getValue("C" + indexAccountNumber, "C" + indexAccountNumber) != null)
                _compNumber = _excelBook.getValue("C" + indexAccountNumber, "C" + indexAccountNumber).ToString().Replace("\n", "");
            if (_excelBook.getValue("D" + indexAccountNumber, "D" + indexAccountNumber) != null)
                _compNumber = _excelBook.getValue("D" + indexAccountNumber, "D" + indexAccountNumber).ToString().Replace("\n", "");
            if (_excelBook.getValue("E" + indexAccountNumber, "E" + indexAccountNumber) != null)
                _compNumber = _excelBook.getValue("E" + indexAccountNumber, "E" + indexAccountNumber).ToString().Replace("\n", "");
            if (_excelBook.getValue("F" + indexAccountNumber, "F" + indexAccountNumber) != null)
                _compNumber = _excelBook.getValue("F" + indexAccountNumber, "F" + indexAccountNumber).ToString().Replace("\n", "");

            if (_day == "")
            {
                if (_excelBook.getValue("B" + i, "B" + i) != null)
                    _date = _excelBook.getValue("B" + i, "B" + i).ToString().Replace("\n", "");
                if (_excelBook.getValue("C" + i, "C" + i) != null)
                    _date = _excelBook.getValue("C" + i, "C" + i).ToString().Replace("\n", "");
                if (_excelBook.getValue("D" + i, "D" + i) != null)
                    _date = _excelBook.getValue("D" + i, "D" + i).ToString().Replace("\n", "");
                if (_excelBook.getValue("E" + i, "E" + i) != null)
                    _date = _excelBook.getValue("E" + i, "E" + i).ToString().Replace("\n", "");
                if (_excelBook.getValue("F" + i, "F" + i) != null)
                    _date = _excelBook.getValue("F" + i, "F" + i).ToString().Replace("\n", "");

                _day = _date.Split(' ')[0];

                if ((Convert.ToInt32(_day) / 10) == 0)
                    _day = "0" + _day;

                _month = _date.Split(' ')[1];
                _year = _date.Substring(_date.Length - 7, 4);
            }
        }

        public void ReadTails(int indexLast)
        {
            int i = indexLast;

            int max = i + 10;

            while (i < max)
            {
                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Входящий остаток"))
                {
                    _currentCell = "D" + i;
                    string sum = _excelBook.getValue("D" + i, "D" + i).ToString().Replace("\n", "");
                    _incomeTail = sum.Substring(0, sum.Length - 4).Replace(" ", "");
                }

                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Исходящий остаток"))
                {
                    _currentCell = "D" + i;
                    string sum = _excelBook.getValue("D" + i, "D" + i).ToString().Replace("\n", "");
                    _outcomeTail = sum.Substring(0, sum.Length - 4).Replace(" ", "");
                }

                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Итого оборотов"))
                {
                    _currentCell = "C" + i;
                    string sum = _excelBook.getValue("C" + i, "C" + i).ToString();
                    _debetTotal = FormatString(sum);

                    _currentCell = "D" + i;
                    sum = _excelBook.getValue("D" + i, "D" + i).ToString();
                    _creditTotal = FormatString(sum);
                }

                i++;
            }
        }

        private string FormatString(string str)
        {
            string result = "";

            foreach (char c in str)
            {
                if (char.IsDigit(c))
                    result += c;
            }

            return result;
        }

        private int GetIndexOnFirst()
        {
            int i = 1;
            while (i < 100)
            {
                if (_excelBook.getValue("A" + i, "A" + i) != null)
                {
                    if ((_excelBook.getValue("A" + i, "A" + i).ToString().Replace("\n", "") == "Дата проводки")
                        || (_excelBook.getValue("A" + i, "A" + i).ToString().Replace("\n", "") == "Дата"))
                    {
                        break;
                    }
                }
                i++;
            }

            i += 2;

            return i;
        }

        public bool IsSumEqualsTotal()
        {
            return IsSumDebetEqualsDebetTotal() && IsSumCreditEqualsCreditTotal();
        }

        public bool IsSumDebetEqualsDebetTotal()
        {
            return ComparisonStrings(debet.Sum, _debetTotal);
        }

        private bool IsSumCreditEqualsCreditTotal()
        {            
            return ComparisonStrings(credit.Sum, _creditTotal);
        }

        private bool ComparisonStrings(string str1, string str2)
        {
            if (str1.Length > str2.Length)
                str1 = str1.Substring(0, str2.Length);
            else if (str1.Length < str2.Length)
                str2 = str2.Substring(0, str1.Length);

            return str1 == str2;
        }
    }
}
