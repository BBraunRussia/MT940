using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MT940
{
    public class File1COpening
    {
        private Settings _settings;

        public File1COpening(Settings settings)
        {
            _settings = settings;
        }

        public string GetFileName()
        {
            if (_settings.Number == "0")
            {
                throw new Exception("Не удалось найти файл с параметрами. В меню Файл->параметры задайте уникальный идентификатор.");
            }
            else
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel files (*.xls)|*.xls";
                ofd.RestoreDirectory = false;
                ofd.Multiselect = false;

                if (ofd.ShowDialog() == DialogResult.OK)
                    return ofd.FileName;
                else
                    throw new Exception("Файл не выбран");
            }
        }
    }
}
