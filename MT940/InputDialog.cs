using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MT940
{
    public partial class InputDialog : Form
    {
        private Invoice _invoice;

        public InputDialog()
        {
            InitializeComponent();

            _invoice = Invoice.GetUniqueInstance();
            input.Text = _invoice.Number;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _invoice.Number = input.Text;
            _invoice.IsRub = radioButton1.Checked;
        }
    }
}
