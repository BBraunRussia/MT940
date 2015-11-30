using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MT940
{
    public partial class fParam : Form
    {
        public fParam()
        {
            InitializeComponent();

            LoadSettings();
        }

        private void LoadSettings()
        {
            if (!File.Exists("settings.ini"))
            {
                return;
            }
            StreamReader sr = new StreamReader("settings.ini");
            string str;
            try
            {
                str = sr.ReadLine();
            }
            finally
            {
                sr.Close();
            }

            str = str.Split('=')[1].Trim();

            int n;

            int.TryParse(str, out n);
            tbN.Text = n.ToString();
        }

        private void SetSettings()
        {
            StreamWriter sw = new StreamWriter("settings.ini", false);
            try
            {
                sw.WriteLine("N = " + tbN.Text.ToString());
            }
            finally
            {
                sw.Close();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SetSettings();
        }
    }
}
