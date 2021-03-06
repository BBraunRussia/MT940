﻿using System;
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
            Converter();
        }

        private void Converter()
        {
            try
            {
                using (ExcelDoc excelBook = new ExcelDoc(FileSberbankOpening.GetFileName()))
                {
                    using (FileTxt fileTxt = new FileTxt())
                    {
                        FileSberbank fileSberbank = new FileSberbank(excelBook);

                        InputDialog id = new InputDialog();

                        if (id.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            fileSberbank.Read();

                            fileSberbank.IsSumDebetEqualsDebetTotal();
                            fileSberbank.IsSumCreditEqualsCreditTotal();

                            fileTxt.Init(fileSberbank, excelBook);
                            fileTxt.WriteBody(FileTxt.TypeRow.D, fileSberbank.Debet);
                            fileTxt.WriteBody(FileTxt.TypeRow.C, fileSberbank.Credit);

                            fileTxt.WriteBottom();

                            MessageBox.Show("Файл сформирован.", "Завершено", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                        else
                        {
                            MessageBox.Show("Пользователь отказался от ввода номера выписки, дальнейшее формирование файла не возможно", "Формирование файла отмененно",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
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

            Close();
        }
    }
}
