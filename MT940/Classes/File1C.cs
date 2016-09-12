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
        private string _month;
        private string _incomeTail;
        private string _outcomeTail;
        private string _debetTotal;
        private string _creditTotal;

        private Invoice _invoice;

        public DCRows Debet { get; private set; }
        public DCRows Credit { get; private set; }

        public string CompNumber { get; private set; }

        public string Day { get; private set; }
        public string MonthDigit { get { return MyMonth.MonthToDigit(_month); } }
        public string Year { get; private set; }

        public string IncomeTail { get { return (_incomeTail == "0,00") ? "0," : _incomeTail; } }
        public string OutcomeTail { get { return (_outcomeTail == "0,00") ? "0," : _outcomeTail; } }

        public string DateFormated { get { return Year.Substring(2, 2) + MonthDigit + Day; } }
        
        public File1C(ExcelDoc excelBook)
        {
            _invoice = Invoice.GetUniqueInstance();
            _excelBook = excelBook;

            _date = "";
            Day = "";
            _month = "";
            Year = "";
            CompNumber = "";
            _incomeTail = "";

            Debet = new DCRows();
            Credit = new DCRows();
        }
        
        public void Read()
        {
            ReadHeader();

            ReadBody();
        }

        private void ReadHeader()
        {
            currentCell = (_invoice.IsRub) ? "O5" : "L4";
            CompNumber = _excelBook.getValue(currentCell, currentCell).ToString().Replace("\n", "");
                        
            if (Day == "")
            {
                currentCell = (_invoice.IsRub) ? "O7" : "L7";
                _date = _excelBook.getValue(currentCell, currentCell).ToString().Replace("\n", "");
                
                Day = _date.Split(' ')[0];

                if ((Convert.ToInt32(Day) / 10) == 0)
                    Day = "0" + Day;

                _month = _date.Split(' ')[1];
                Year = _date.Substring(_date.Length - 7, 4);
            }
        }
                
        public bool IsSumDebetEqualsDebetTotal()
        {
            if (ComparisonStrings(Debet.Sum, _debetTotal))
                return true;
            else
                throw new OverflowException(string.Concat("Формирование файла отменено, так как сумма по дебету (", Debet.Sum, ") не совпадает с итоговым значением (", _debetTotal, ")." ));
        }

        public bool IsSumCreditEqualsCreditTotal()
        {
            if (ComparisonStrings(Credit.Sum, _creditTotal))
                return true;
            else
                throw new OverflowException(string.Concat("Формирование файла отменено, так как сумма по кредиту (", Credit.Sum, ") не совпадает с итоговым значением (", _creditTotal, ")."));
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
                currentCell = (_invoice.IsRub) ? "Q" + i : "D" + i;
                while (_excelBook.getValue(currentCell, currentCell) != null)
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
                            Debet.Add(dcRow);
                    }
                    else
                    {
                        currentCell = (_invoice.IsRub) ? "P" + i : "O" + i; //Сумма по кредиту
                        dcRow.SetSum(_excelBook.getValue(currentCell, currentCell));

                        if (dcRow.Sum != 0.0)
                            Credit.Add(dcRow);
                    }

                    i++;
                    currentCell = (_invoice.IsRub) ? "Q" + i : "D" + i;
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
                currentCell = (_invoice.IsRub) ? "C" + i : "B" + i;
                if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() == "Входящий остаток"))
                {
                    currentCell = (_invoice.IsRub) ? "N" + i : "N" + i;
                    _incomeTail = FormatTail(_excelBook.getValue(currentCell, currentCell).ToString());
                }

                currentCell = (_invoice.IsRub) ? "C" + i : "B" + i;
                if ((_excelBook.getValue(currentCell, currentCell) != null) && (_excelBook.getValue(currentCell, currentCell).ToString() == "Исходящий остаток"))
                {
                    currentCell = (_invoice.IsRub) ? "N" + i : "N" + i;
                    _outcomeTail = FormatTail(_excelBook.getValue(currentCell, currentCell).ToString());
                }

                currentCell = (_invoice.IsRub) ? "C" + i : "B" + i;
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
