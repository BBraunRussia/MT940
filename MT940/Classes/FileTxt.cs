using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MT940
{
    public class FileTxt : IDisposable
    {
        public enum TypeRow { D, C }

        private StreamWriter _streamWriter;
        private FileSberbank _file1C;
        private ExcelDoc _excelBook;

        private Invoice _invoice;

        public void Init(FileSberbank file1C, ExcelDoc excelBook)
        {
            _file1C = file1C;
            _excelBook = excelBook;

            _invoice = Invoice.GetUniqueInstance();

            _streamWriter = new StreamWriter(Path.GetDirectoryName(_excelBook.FileName) + @"\" + file1C.CompNumber + "_" + file1C.Day + file1C.MonthDigit +
                    file1C.Year.Substring(2, 2) + "_vip.txt", false, Encoding.Unicode);

            WriteHeader();
        }

        private void WriteHeader()
        {
            WriteLine("{1:F01SABRRU2PAXXX0000000000}{2:I940SABRRU2PXXXXN}{4:");

            WriteLine(":21:NONREF");
            WriteLine(string.Concat(":25:", _file1C.CompNumber));

            WriteLine(string.Concat(":28C:", _invoice.GetNumberFormated()));

            WriteLine(string.Concat(":60F:C", _file1C.DateFormated, _invoice.Currency, _file1C.IncomeTail));
        }

        public void WriteBody(TypeRow type, DCRows dcRows)
        {
            foreach (DCRow dcRow in dcRows.List)
            {
                string sum = dcRow.SumString.ToString();
                if (sum.IndexOf(',') == -1)
                    sum += ",";

                string letters = (_invoice.IsRub) ? "NCMI" : "NTRF";

                WriteLine(string.Concat(":61:", _file1C.DateFormated, type.ToString(), sum, letters, dcRow.Number, "//", dcRow.Number));
                WriteLine(string.Concat(":86:/ORDP/", dcRow.Ordp));
                WriteLine(string.Concat("/BENM/", dcRow.Benm));

                foreach (string s1 in dcRow.CommList)
                {
                    WriteLine(s1);
                }
            }
        }

        public void WriteBottom()
        {
            WriteLine(string.Concat(":62F:C", _file1C.DateFormated, _invoice.Currency, _file1C.OutcomeTail));
            WriteLine("-}");
        }

        public void WriteLine(string text)
        {
            _streamWriter.WriteLine(text);
        }

        public void Dispose()
        {
            if (_streamWriter != null)
                _streamWriter.Close();
        }
    }
}
