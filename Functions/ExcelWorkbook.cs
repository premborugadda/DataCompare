using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataCompare.Functions
{
    class ExcelWorkbook
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel._Worksheet xlWorksheet;
        public Excel.Range xlRange;

        public ExcelWorkbook(string workbookPath)
        {
            WorkbookPath = workbookPath;

            xlApp = new Excel.Application();
            xlWorkbook = openWorkbook();
            xlWorksheet = openSheet(1);
            xlRange = getRange();
        }

        public string WorkbookPath { get; }

        // TODO fill openWorkbook()
        public Excel.Workbook openWorkbook()
        {
            return xlApp.Workbooks.Open(WorkbookPath);
        }

        // TODO fill closeWorkbook()
        public void closeWorkbook()
        {
            xlApp.Workbooks.Close();
        }

        // TODO fill openSheet()
        public Excel._Worksheet openSheet(int sheetIndex)
        {
            return xlWorkbook.Sheets[sheetIndex];
        }

        public Excel.Range getRange()
        {
            return xlWorksheet.UsedRange;
        }

        // TODO fill getCellValue()
        public string getCellValue(int i, int j)
        {
            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            {
                Excel.Range range = (xlWorksheet.Cells[i, j] as Excel.Range);
                string cellValue = range.Value.ToString();
                return cellValue;
            }else
            {
                return "";
            }

            
        }

    }
}

