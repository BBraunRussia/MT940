using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace MT940
{
    public class ExcelDoc : OfficeDoc, IDisposable
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkBook;
        private Excel.Worksheet xlSh;

        public string FileName { get { return name; } }

        public ExcelDoc(string name)
            : base(name)
        {
            Init();
        }

        public ExcelDoc()
        {
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlSh = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        private void Init()
        {
            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Open(name, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlSh = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        public void setValue(int rowIndex, int columnIndex, string value)
        {
            xlSh.Cells[rowIndex, columnIndex] = value;
        }

        public object getValue(string rowCell, string columnCell)
        {
            return xlSh.get_Range(rowCell, columnCell).Value2;
        }

        public void Show()
        {
            xlSh.Columns.AutoFit();

            xlApp.Visible = true;
        }

        public void Dispose()
        {
            xlApp.DisplayAlerts = false;
            xlApp.Quit();

            releaseObject(xlSh);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        internal void Print()
        {
            object misValue = System.Reflection.Missing.Value;

            xlSh.Columns.AutoFit();

            xlSh.PrintOut(1, 1, 1, false, misValue, misValue, misValue, misValue);

            Dispose();
        }
    }

    public class OfficeDoc
    {
        protected string name;

        protected OfficeDoc()
        {
        }

        protected OfficeDoc(string name)
        {
            this.name = name;
        }

        protected void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
