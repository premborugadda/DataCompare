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

        public List<Budget> GetValues (string BudgetSheetLoc)
        {
            List<Budget> records = new List<Budget>();
            wb = new ExcelWorkbook(BudgetSheetLoc);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;
            var colCount = wb.xlRange.Columns.Count;

            for (int i = 3; i <= 14; i++)
            {
                for (int j=12; j <= 21; j++)
                {

                
                Budget mine = new Budget();

                mine.mineName = wb.getCellValue(j, 2);
                mine.budgetMonth = wb.getCellValue(8, i);

                double line1 = 0;
                Double.TryParse(wb.getCellValue(j, i), out line1);
                mine.budgetValue = line1*1000;

                //double line2 = 0;
                //Double.TryParse(wb.getCellValue(i, 4), out line2);
                //mine.FebBudget = line2;

                //double line3 = 0;
                //Double.TryParse(wb.getCellValue(i, 5), out line3);
                //mine.MarBudget = line3;

                //double line4 = 0;
                //Double.TryParse(wb.getCellValue(i, 6), out line4);
                //mine.AprBudget = line4;

                //double line5 = 0;
                //Double.TryParse(wb.getCellValue(i, 7), out line5);
                //mine.MayBudget = line5;

                //double line6 = 0;
                //Double.TryParse(wb.getCellValue(i, 8), out line6);
                //mine.JunBudget = line6;

                //double line7 = 0;
                //Double.TryParse(wb.getCellValue(i, 9), out line7);
                //mine.JulBudget = line7;

                //double line8 = 0;
                //Double.TryParse(wb.getCellValue(i, 10), out line8);
                //mine.AugBudget = line8;

                //double line9 = 0;
                //Double.TryParse(wb.getCellValue(i, 11), out line9);
                //mine.SepBudget = line9;

                //double line10 = 0;
                //Double.TryParse(wb.getCellValue(i, 12), out line10);
                //mine.OctBudget = line10;

                //double line11 = 0;
                //Double.TryParse(wb.getCellValue(i, 13), out line11);
                //mine.NovBudget = line11;

                //double line12 = 0;
                //Double.TryParse(wb.getCellValue(i, 14), out line12);
                //mine.DecBudget = line12;

                //double line13 = 0;
                //Double.TryParse(wb.getCellValue(i, 15), out line13);
                //mine.YearBudget = line13;

                records.Add(mine);
                }

            }
            return records;
        }
    }
}
