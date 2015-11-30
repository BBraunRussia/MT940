using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MT940
{
    public class DCRow
    {
        private double _sum;
        private string _number;
        private string _ordp;
        private string _benm;
        private List<string> _commList;

        public double Sum { get { return _sum; } }
        public string SumString { get { return _sum.ToString(); } }
        public string Number { get { return _number; } }
        public string Ordp { get { return _ordp; } }
        public string Benm { get { return _benm; } }

        public List<string> CommList { get { return _commList; } }

        public DCRow()
        {
            _commList = new List<string>();
        }

        public void SetSum(object value)
        {
            if (value == null)
                return;

            double.TryParse(value.ToString().Replace(" ", "").Replace(".", ","), out _sum);
        }

        public void SetNumber(object value)
        {
            if (value == null)
                return;

            _number = value.ToString();
        }

        public void SetOrdp(object value)
        {
            if (value == null)
                return;

            _ordp = value.ToString();
        }

        public void SetOrdpWithoutDigit(object value)
        {
            SetOrdp(value);

            for (int j = 0; j < _ordp.Length; j++)
            {
                if (_ordp[j] == ' ')
                {
                    _ordp = _ordp.Substring(j + 1);
                    break;
                }
            }
        }

        public void SetBenm(object value)
        {
            if (value == null)
                return;

            _benm = value.ToString();
        }

        public void SetBenmWithoutDigit(object value)
        {
            SetBenm(value);

            for (int j = 0; j < _benm.Length; j++)
            {
                if (_benm[j] == ' ')
                {
                    _benm = _benm.Substring(j + 1);
                    break;
                }
            }
        }

        public void SetCom(object value)
        {
            if (value == null)
                return;

            CreateCommList(value.ToString());
        }

        private void CreateCommList(string value)
        {
            value = GetSubString(value);

            string[] mas;
            mas = value.Split(' ');
            string temp = "/REMI/";
            foreach (string s in mas)
            {
                if ((temp.Length + s.Length) < 66)
                {
                    temp += (temp == "/REMI/") ? s : " " + s;
                }
                else
                {
                    _commList.Add(temp + "\n");
                    temp = "//" + s;
                }
            }
            _commList.Add(temp + "\n");
        }

        private string GetSubString(string value)
        {
            return (value.Length > 210) ? value.Substring(0, 210) : value;
        }
    }
}
