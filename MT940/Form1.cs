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
    public partial class Form1 : Form
    {
        Settings settings;

        public Form1()
        {
            InitializeComponent();

            ReadSetting();
        }

        private void ReadSetting()
        {
            try
            {
                settings = new Settings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
                
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fParam fp = new fParam();
            if (fp.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                settings.Read();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fHelp fh = new fHelp();
            fh.ShowDialog();
        }

        private void btnCreateFile_Click(object sender, EventArgs e)
        {
            try
            {
                File1COpening file = new File1COpening(settings);

                Converter(file.GetFileName());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Converter(string fileName)
        {
            ExcelDoc excelBook = new ExcelDoc(fileName);

            File1C file1C = new File1C(excelBook);
            FileTxt fileTxt = new FileTxt();

            try
            {
                Invoice invoice = new Invoice();
                InputDialog id = new InputDialog(invoice);

                if (id.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    file1C.Read();

                    file1C.IsSumDebetEqualsDebetTotal();
                    file1C.IsSumCreditEqualsCreditTotal();

                    fileTxt.Init(file1C, settings, excelBook, invoice);
                    fileTxt.WriteBody(FileTxt.TypeRow.D, file1C.Debet);
                    fileTxt.WriteBody(FileTxt.TypeRow.C, file1C.Credit);

                    fileTxt.WriteBottom();

                    MessageBox.Show("Файл сформирован.", "Завершено", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("Пользователь отказался от ввода номера выписки, дальнейшее формирование файла не возможно", "Формирование файла отмененно",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.Message, "Формирование файла отмененно", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (OverflowException ex)
            {
                MessageBox.Show(ex.Message, "Формирование файла отмененно", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                /*
            catch
            {
                if (file1C.CurrentCell != "")
                {
                    MessageBox.Show("Ошибка при обработке файла. Проверьте ячейку " + file1C.CurrentCell, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Ошибка при обработке файла", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
                 * */
            finally
            {
                excelBook.Dispose();
                fileTxt.Dispose();
            }

            Close();
        }
    }
}
