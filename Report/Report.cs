using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using DataCompare.Functions;
using Microsoft.Office.Interop.Excel;

namespace DataCompare.Report
{
    class Report
    {
        private ExcelWorkbook wb;
        private List<Result> resMines;
        private Result r;

        public Report()
        {            
            this.resMines = new List<Result>();            
        }

        public void createResMine(string mineName, int indexColumn)
        {
            try
            {
                r = new Result(mineName);
                
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        public Result getResMine(string mineName)
        {
            return resMines.Find(mine => mine.ResMineName == mineName);
        }

        public void ReportDouble(string SheetLoc, int Sheetindex, int i, int j, double? Value1, string Value2)
        {
            wb = new ExcelWorkbook(SheetLoc);
            var xlWorksheet = wb.openSheet(Sheetindex);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;
            var colCount = wb.xlRange.Columns.Count;

            Excel.Range range = (xlWorksheet.Cells[i, j] as Excel.Range);
            if (Value1 != 0)
            {
                range.Value = Value1;
            }
            else if (Value2 != "")
            {
                range.Value = Value2;
            }
            else
            {
                range.Value = null;
            }

            //xlWorksheet.SaveAs(SheetLoc, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

            xlWorksheet.SaveAs(SheetLoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            wb.closeWorkbook();
            ReleaseObject(xlWorksheet);
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        public void writeData(string SheetLoc, string mineName, List<Result> resMines, string toDate)
        {
            wb = new ExcelWorkbook(SheetLoc);
            var xlWorksheet = wb.openSheet(2);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;
            var colCount = wb.xlRange.Columns.Count;
            string[] mineNames = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "OVOID MINE", "THOMPSON MINE", "TOTTEN MINE" };
            //foreach (string minename in mineNames)
            //{
            //    resMines = this.getResMine(mineName);
            //}
            int iCount = resMines.Count(); 
            int intRow = 3;

            xlWorksheet.Cells[1, 4] = toDate;
            xlWorksheet.Cells[1, 6] = DateTime.Today;

            for (int i = 0; i < iCount; i++)
            {
                xlWorksheet.Cells[intRow, 9] = resMines[i].Actual.AppDay;
                xlWorksheet.Cells[intRow, 10] = resMines[i].Actual.HanaDay;
                xlWorksheet.Cells[intRow, 11] = resMines[i].Actual.NaidDay;

                xlWorksheet.Cells[intRow, 12] = resMines[i].Actual.AppMon;
                xlWorksheet.Cells[intRow, 13] = resMines[i].Actual.HanaMon;
                xlWorksheet.Cells[intRow, 14] = resMines[i].Actual.NaidMon;

                xlWorksheet.Cells[intRow, 15] = resMines[i].Actual.AppYear;
                xlWorksheet.Cells[intRow, 16] = resMines[i].Actual.HanaYear;
                xlWorksheet.Cells[intRow, 17] = resMines[i].Actual.NaidYear;
                
                xlWorksheet.Cells[intRow, 18] = resMines[i].Budget.AppDay;                
                xlWorksheet.Cells[intRow, 19] = resMines[i].Budget.HanaDay;
                xlWorksheet.Cells[intRow, 20] = resMines[i].Budget.NaidDay;
                
                xlWorksheet.Cells[intRow, 21] = resMines[i].Budget.AppMon;
                xlWorksheet.Cells[intRow, 22] = resMines[i].Budget.HanaMon;
                xlWorksheet.Cells[intRow, 23] = resMines[i].Budget.NaidMon;
                
                xlWorksheet.Cells[intRow, 24] = resMines[i].Budget.AppYear + resMines[i].Budget.AppMon;
                xlWorksheet.Cells[intRow, 25] = resMines[i].Budget.HanaYear;
                xlWorksheet.Cells[intRow, 26] = resMines[i].Budget.AppYear;
                xlWorksheet.Cells[intRow, 27] = resMines[i].Budget.NaidYear;                

                intRow++;


            }


            xlWorksheet.SaveAs(SheetLoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            wb.closeWorkbook();
            ReleaseObject(xlWorksheet);
        }

    }

}
