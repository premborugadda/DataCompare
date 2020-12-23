using DataCompare.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.BudgetParser
{
    public class dataParse
    {
    
    private ExcelWorkbook wb;

        public List<Budget> GetValues (string BudgetSheetLoc, List<string> mineNames, int[] reqRows)
        {
            List<Budget> records = new List<Budget>();
            wb = new ExcelWorkbook(BudgetSheetLoc);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;
            var colCount = wb.xlRange.Columns.Count;
            //int[] reqRows = {15, 12, 13, 14, 33, 27, 19};


            for (int i = 3; i <= 14; i++)
            {
                for (int j=0; j <= reqRows.Length-1; j++)
                {
                
                Budget mine = new Budget();
                mine.mineName = mineNames[j]; // wb.getCellValue(j, 2).ToLower();// + " MINE";
                mine.budgetMonth = wb.getCellValue(8, i);

                double line1 = 0;
                Double.TryParse(wb.getCellValue(reqRows[j], i), out line1);
                mine.budgetValue = line1*1000;  
                records.Add(mine);
                }

            }
            return records;
        }
    }
}
