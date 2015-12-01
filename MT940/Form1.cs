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
        public Form1()
        {
            InitializeComponent();
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
            File1COpening file = new File1COpening();

            Converter(file.GetFileName());
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

                    fileTxt.Init(file1C, excelBook, invoice);
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
            catch (NotImplementedException ex)
            {
                MessageBox.Show(ex.Message, "Формирование файла отмененно", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                excelBook.Dispose();
                fileTxt.Dispose();
            }

            Close();
        }
    }
}
