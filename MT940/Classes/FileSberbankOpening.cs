using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MT940
{
    public static class FileSberbankOpening
    {
        public static string GetFileName()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xls)|*.xls";
            ofd.RestoreDirectory = false;
            ofd.Multiselect = false;

            if (ofd.ShowDialog() == DialogResult.OK)
                return ofd.FileName;
            else
                throw new NullReferenceException("Файл не выбран");
        }
    }
}
