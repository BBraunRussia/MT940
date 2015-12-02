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
        private File1C _file1C;
        private ExcelDoc _excelBook;

        public void Init(File1C file1C, ExcelDoc excelBook, Invoice invoice)
        {
            _file1C = file1C;
            _excelBook = excelBook;

            _streamWriter = new StreamWriter(Path.GetDirectoryName(_excelBook.FileName) + @"\" + file1C.CompNumber + "_" + file1C.Day + file1C.MonthDigit +
                    file1C.Year.Substring(2, 2) + "_vip.txt", false, Encoding.Unicode);

            WriteHeader(invoice);
        }

        private void WriteHeader(Invoice invoice)
        {
            Settings settings = Settings.GetUniqueInstance();

            WriteLine("{1:F01SABRRU2PAXXX0000000000}{2:I940SABRRU2PXXXXN}{4:");
            WriteLine(":20:+5500" + _file1C.DateFormated + "0" + settings.Number);

            settings.Save();

            WriteLine(":21:NONREF");
            WriteLine(":25:" + _file1C.CompNumber);

            WriteLine(":28C:" + invoice.GetNumberFormated());
            WriteLine(":60F:C" + _file1C.DateFormated + "RUB" + _file1C.IncomeTail);
        }

        public void WriteBody(TypeRow type, DCRows dcRows)
        {
            foreach (DCRow dcRow in dcRows.List)
            {
                string temp = dcRow.SumString.ToString();
                if (temp.IndexOf(',') == -1)
                    temp += ",";
                WriteLine(":61:" + _file1C.DateFormated + type.ToString() + temp + "NCMI" + dcRow.Number + "//" + dcRow.Number);
                WriteLine(":86:/ORDP/" + dcRow.Ordp);
                WriteLine("/BENM/" + dcRow.Benm);

                foreach (string s1 in dcRow.CommList)
                {
                    WriteLine(s1);
                }
            }
        }

        public void WriteBottom()
        {
            WriteLine(":62F:C" + _file1C.DateFormated + "RUB" + _file1C.OutcomeTail);
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
