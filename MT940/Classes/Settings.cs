using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MT940
{
    public class Settings
    {
        private int _number;

        public string Number
        {
            get { return _number.ToString(); }
        }

        public Settings()
        {
            Read();
        }

        public void Read()
        {
            if (File.Exists("settings.ini"))
            {
                using (StreamReader sr = new StreamReader("settings.ini"))
                {
                    try
                    {
                        string str;
                        str = sr.ReadLine();
                        str = str.Split('=')[1].Trim();
                        int.TryParse(str, out _number);
                    }
                    catch { }
                }

                if (_number == 0)
                {
                    throw new Exception("Не удалось распознать уникальный идентификатор. В меню Файл->параметры задайте новое значение уникального идентификатора.");
                }
            }
            else
            {
                throw new Exception("Не удалось найти файл с параметрами. В меню Файл->параметры задайте новое значение уникального идентификатора.");
            }
        }

        public void Save()
        {
            _number++;
            using (StreamWriter sw = new StreamWriter("settings.ini"))
            {
                sw.WriteLine("N = " + _number.ToString());
            }
        }
    }
}
