﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MT940
{
    public class Settings
    {
        private int _numberBBraun;
        private int _numberGematek;
        private string _nameBBraun;
        private string _nameGematek;
        private bool _isBBraunFile;

        private string _comment;

        public string Number { get { return (_isBBraunFile) ? _numberBBraun.ToString() : _numberGematek.ToString(); } }
        public bool IsBBraunFile { set { _isBBraunFile = value; } }

        private static Settings _uniqueInstance;

        private Settings()
        {
            Read();
        }

        public static Settings GetUniqueInstance()
        {
            if (_uniqueInstance == null)
                _uniqueInstance = new Settings();

            return _uniqueInstance;
        }

        private void Read()
        {
            if (File.Exists("settings.ini"))
            {
                using (StreamReader sr = new StreamReader("settings.ini"))
                {
                    string str;
                    str = sr.ReadLine(); //примечание
                    _comment = str;

                    str = sr.ReadLine();
                    string[] line = str.Split('=');
                    _nameBBraun = line[0].Trim();
                    int.TryParse(line[1].Trim(), out _numberBBraun);

                    str = sr.ReadLine();
                    line = str.Split('=');
                    _nameGematek = line[0].Trim();
                    int.TryParse(line[1].Trim(), out _numberGematek);
                }

                if (_numberBBraun == 0)
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
            if (_isBBraunFile)
                _numberBBraun++;
            else
                _numberGematek++;

            using (StreamWriter sw = new StreamWriter("settings.ini"))
            {
                sw.WriteLine(_comment);
                sw.WriteLine(string.Concat(_nameBBraun, " = ", _numberBBraun.ToString()));
                sw.WriteLine(string.Concat(_nameGematek, " = ", _numberGematek.ToString()));
            }
        }
    }
}
