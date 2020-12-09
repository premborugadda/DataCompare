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

    }

}




