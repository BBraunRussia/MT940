using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public class File1C
    {
        public static string currentCell;

        private ExcelDoc _excelBook;

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

        private Invoice _invoice;

        public DCRows Debet { get { return debet; } }
        public DCRows Credit { get { return credit; } }

        public string CompNumber { get { return _compNumber; } }

        public string Day { get { return _day; } }
        public string MonthDigit { get { return MyMonth.MonthToDigit(_month); } }
        public string Year { get { return _year; } }

        public string IncomeTail { get { return (_incomeTail == "0,00") ? "0," : _incomeTail; } }
        public string OutcomeTail { get { return (_outcomeTail == "0,00") ? "0," : _outcomeTail; } }

        public string DateFormated { get { return _year.Substring(2, 2) + MonthDigit + _day; } }
        
        public File1C(ExcelDoc excelBook)
        {
            _invoice = Invoice.GetUniqueInstance();
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
            currentCell = (_invoice.IsRub) ? "O5" : "L4";
            _compNumber = _excelBook.getValue(currentCell, currentCell).ToString().Replace("\n", "");
                        
            if (_day == "")
            {
                currentCell = (_invoice.IsRub) ? "O8" : "L7";
                _date = _excelBook.getValue(currentCell, currentCell).ToString().Replace("\n", "");
                
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
            int i = (_invoice.IsRub) ? 12 : 11; //первая строка с данными
            int incement = (_invoice.IsRub) ? 7 : 4; //первая строка с данными
            int readBlocks = 0;
            int countBlocks = GetCountBlocks();

            while (readBlocks < countBlocks)
            {
                while (_excelBook.getValue("F" + i, "F" + i) != null)
                {
                    DCRow dcRow = new DCRow();

                    currentCell = (_invoice.IsRub) ? "Q" + i : "D" + i; //№ документа
                    dcRow.SetNumber(_excelBook.getValue(currentCell, currentCell));

                    currentCell = (_invoice.IsRub) ? "F" + i : "E" + i; //Счёт дебет
                    dcRow.SetOrdp(_excelBook.getValue(currentCell, currentCell));

                    currentCell = (_invoice.IsRub) ? "J" + i : "G" + i; //Счёт кредит
                    dcRow.SetBenm(_excelBook.getValue(currentCell, currentCell));

                    currentCell = (_invoice.IsRub) ? "Y" + i : "Y" + i; //Назначение платежа
                    dcRow.SetCom(_excelBook.getValue(currentCell, currentCell));

                    currentCell = (_invoice.IsRub) ? "L" + i : "I" + i; //Сумма по дебету
                    if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() != string.Empty))
                    {
                        dcRow.SetSum(_excelBook.getValue(currentCell, currentCell));

                        if (dcRow.Sum != 0.0)
                            debet.Add(dcRow);
                    }
                    else
                    {
                        currentCell = (_invoice.IsRub) ? "P" + i : "O" + i; //Сумма по кредиту
                        dcRow.SetSum(_excelBook.getValue(currentCell, currentCell));

                        if (dcRow.Sum != 0.0)
                            credit.Add(dcRow);
                    }

                    i++;
                }

                i += incement;
                readBlocks++;
            }

            ReadTails(i - incement);
        }

        public void ReadTails(int i)
        {
            int max = i + 20;

            while (i < max)
            {
                currentCell = (_invoice.IsRub) ? "C" + i : "C" + i;
                if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() == "Входящий остаток"))
                {
                    currentCell = (_invoice.IsRub) ? "N" + i : "N" + i;
                    _incomeTail = FormatTail(_excelBook.getValue(currentCell, currentCell).ToString());
                }

                currentCell = (_invoice.IsRub) ? "C" + i : "C" + i;
                if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() == "Исходящий остаток"))
                {
                    currentCell = (_invoice.IsRub) ? "N" + i : "N" + i;
                    _outcomeTail = FormatTail(_excelBook.getValue(currentCell, currentCell).ToString());
                }

                currentCell = (_invoice.IsRub) ? "C" + i : "C" + i;
                if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() == "Итого оборотов"))
                {
                    currentCell = (_invoice.IsRub) ? "I" + i : "F" + i;
                    string formatTotal = FormatTotal(_excelBook.getValue(currentCell, currentCell).ToString());
                    _debetTotal = FormatString(formatTotal);

                    currentCell = (_invoice.IsRub) ? "N" + i : "N" + i;
                    formatTotal = FormatTotal(_excelBook.getValue(currentCell, currentCell).ToString());
                    _creditTotal = FormatString(formatTotal);
                }

                i++;
            }

            Validate("Входящий остаток", _incomeTail);
            Validate("Исходящий остаток", _outcomeTail);
        }

        private void Validate(string valueName, string value)
        {
            if (value == null)
                throw new NullReferenceException("Не заполнено поле " + valueName);
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
            int spaceCount = (_invoice.IsRub) ? 5 : 2;

            int i = 11;
            int countNull = 0;
            int countBlocks = 1;

            while (countNull < 6)
            {
                currentCell = (_invoice.IsRub) ? "F" + i : "D" + i;

                if (_excelBook.getValue(currentCell, currentCell) == null)
                    countNull++;
                else
                {
                    if (countNull == spaceCount)
                        countBlocks++;
                    countNull = 0;
                }

                i++;
            }

            return countBlocks;
        }

        private string FormatTail(string sum)
        {
            sum = (_invoice.IsRub) ? sum.Substring(0, sum.Length - 3) : sum.Substring(0, sum.Length - 13).Split('/')[1];
            return DeleteSplits(sum);
        }

        private string FormatTotal(string sum)
        {
            sum = (_invoice.IsRub) ? sum : sum.Substring(0, sum.Length - 9).Split('/')[1];
            return DeleteSplits(sum);
        }

        public static string DeleteSplits(string value)
        {
            return (value.Contains('.')) ? value.Replace(",", " ").Replace(".", ",").Replace(" ", "") : value.Replace(" ", "");
        }
    }
}
