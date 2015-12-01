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

            ReadBody();
        }

        private void ReadHeader()
        {
            string orgName = _excelBook.getValue("N5", "N5").ToString().Replace("\n", "");
            Settings settings = Settings.GetUniqueInstance();

            if (orgName == "ООО \"Б.Браун Медикал\"")
                settings.IsBBraunFile = true;
            else if (orgName == "ООО \"Гематек\"")
                settings.IsBBraunFile = false;
            else
                throw new NotImplementedException(string.Concat("Для организации ", orgName, " не реализован конвертер"));

            _compNumber = _excelBook.getValue("N4", "N4").ToString().Replace("\n", "");
                        
            if (_day == "")
            {
                _date = _excelBook.getValue("N7", "N7").ToString().Replace("\n", "");
                
                _day = _date.Split(' ')[0];

                if ((Convert.ToInt32(_day) / 10) == 0)
                    _day = "0" + _day;

                _month = _date.Split(' ')[1];
                _year = _date.Substring(_date.Length - 7, 4);
            }
        }
                
        public bool IsSumDebetEqualsDebetTotal()
        {
            if (ComparisonStrings(debet.Sum, _debetTotal))
                return true;
            else
                throw new OverflowException(string.Concat("Формирование файла отменено, так как сумма по дебету (", debet.Sum, ") не совпадает с итоговым значением (", _debetTotal, ")." ));
        }

        public bool IsSumCreditEqualsCreditTotal()
        {
            if (ComparisonStrings(credit.Sum, _creditTotal))
                return true;
            else
                throw new OverflowException(string.Concat("Формирование файла отменено, так как сумма по кредиту (", credit.Sum, ") не совпадает с итоговым значением (", _creditTotal, ")."));
        }

        private bool ComparisonStrings(string str1, string str2)
        {
            if (str1.Length > str2.Length)
                str1 = str1.Substring(0, str2.Length);
            else if (str1.Length < str2.Length)
                str2 = str2.Substring(0, str1.Length);

            return str1 == str2;
        }

        private void ReadBody()
        {
            int i = 11;
            int readBlocks = 0;
            int countBlocks = GetCountBlocks();

            while (readBlocks < countBlocks)
            {
                while (_excelBook.getValue("O" + i, "O" + i) != null)
                {
                    DCRow dcRow = new DCRow();

                    CurrentCell = "O" + i;
                    dcRow.SetNumber(_excelBook.getValue("O" + i, "O" + i));

                    CurrentCell = "E" + i;
                    dcRow.SetOrdp(_excelBook.getValue("E" + i, "E" + i));

                    CurrentCell = "H" + i;
                    dcRow.SetBenm(_excelBook.getValue("H" + i, "H" + i));

                    CurrentCell = "V" + i;
                    dcRow.SetCom(_excelBook.getValue("V" + i, "V" + i));

                    CurrentCell = "J" + i;
                    if ((_excelBook.getValue("J" + i, "J" + i) != null) && (_excelBook.getValue("J" + i, "J" + i).ToString() != string.Empty))
                    {
                        dcRow.SetSum(_excelBook.getValue("J" + i, "J" + i));

                        debet.Add(dcRow);
                    }
                    else
                    {
                        CurrentCell = "M" + i;
                        dcRow.SetSum(_excelBook.getValue("M" + i, "M" + i));

                        credit.Add(dcRow);
                    }

                    i++;
                }

                i += 4;
                readBlocks++;
            }

            ReadTails(i);
        }

        public void ReadTails(int i)
        {
            int max = i + 10;

            while (i < max)
            {
                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Входящий остаток"))
                {
                    _currentCell = "L" + i;
                    string sum = _excelBook.getValue("L" + i, "L" + i).ToString().Replace("\n", "");
                    _incomeTail = sum.Substring(0, sum.Length - 4).Replace(" ", "");
                }

                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Исходящий остаток"))
                {
                    _currentCell = "L" + i;
                    if ((_excelBook.getValue("L" + i, "L" + i) == null) || (_excelBook.getValue("L" + i, "L" + i).ToString() == string.Empty))
                        throw new NullReferenceException("Нет данных в ячейки с исходящим остатком");

                    string sum = _excelBook.getValue("L" + i, "L" + i).ToString().Replace("\n", "");
                    
                    _outcomeTail = sum.Substring(0, sum.Length - 4).Replace(" ", "");
                }

                if ((_excelBook.getValue("B" + i, "B" + i) != null) && (_excelBook.getValue("B" + i, "B" + i).ToString() == "Итого оборотов"))
                {
                    _currentCell = "G" + i;
                    string sum = _excelBook.getValue("G" + i, "G" + i).ToString();
                    _debetTotal = FormatString(sum);

                    _currentCell = "L" + i;
                    sum = _excelBook.getValue("L" + i, "L" + i).ToString();
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

        private int GetCountBlocks()
        {
            int i = 11;
            int countNull = 0;
            int countBlocks = 1;

            while (countNull < 3)
            {
                if ((_excelBook.getValue("O" + i, "O" + i) == null) || (_excelBook.getValue("O" + i, "O" + i) == null))
                    countNull++;
                else
                {
                    if (countNull == 2)
                        countBlocks++;
                    countNull = 0;
                }

                i++;
            }

            return countBlocks;
        }
    }
}
