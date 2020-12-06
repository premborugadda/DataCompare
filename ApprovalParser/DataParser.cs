using DataCompare.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.ApprovalParser
{
    public class DataParser
    {
        private ExcelWorkbook wb;

        public List<Approval> ReadValues (string AppsheetLocation)
        {
            List<Approval> records = new List<Approval>();
            wb = new ExcelWorkbook(AppsheetLocation);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;
            var colCount = wb.xlRange.Columns.Count;

            for (int i = 2; i < rowCount; i++)
            {   
                Approval mine = new Approval();

                //mine.CalendarDate = System.DateTime.ParseExact(wb.getCellValue(i, 1), "yyyy-MM-dd", new CultureInfo("en-us"));
                mine.CalendarDate = Convert.ToDateTime(wb.getCellValue(i, 1));
                    
                double line1 = 0;
                Double.TryParse(wb.getCellValue(i, 2), out line1);
                mine.DimOperationalMeasureKey = line1;

                double line2 = 0;
                Double.TryParse(wb.getCellValue(i, 3), out line2);
                mine.RevisedVoulmeDry = line2;

                double line3 = 0;
                Double.TryParse(wb.getCellValue(i, 4), out line3);
                mine.PrelinVolume = line3;

                double line4 = 0;
                Double.TryParse(wb.getCellValue(i, 5), out line4);
                mine.PrelinDrynessFactor = line4;

                double line5 = 0;
                Double.TryParse(wb.getCellValue(i, 6), out line5);
                mine.RevisedVoulmeWet = line5;

                double line6 = 0;
                Double.TryParse(wb.getCellValue(i, 7), out line6);
                mine.RevisedDrynessFactor = line6;

                double line7 = 0;
                Double.TryParse(wb.getCellValue(i, 8), out line7);
                mine.CrushedOreAdjusted = line7;

                mine.Comments = wb.getCellValue(i, 9);
                //mine.CreatedDate = System.DateTime.ParseExact(wb.getCellValue(i, 10), "yyyy-MM-dd", new CultureInfo("en-us"));
                mine.CreatedDate = Convert.ToDateTime(wb.getCellValue(i, 10));

                double line10 = 0;
                Double.TryParse(wb.getCellValue(i, 11), out line10);
                mine.DimDateKey = line10;

                //mine.UpdatedDate = System.DateTime.ParseExact(wb.getCellValue(i, 12), "yyyy-MM-dd", new CultureInfo("en-us"));
                //mine.UpdatedDate = Convert.ToDateTime(wb.getCellValue(i, 12));
                mine.UpdatedBy = wb.getCellValue(i, 13);
                mine.CreatedBy = wb.getCellValue(i, 14);
                mine.UpdatedByEmailId = wb.getCellValue(i, 15);

                double line15 = 0;
                Double.TryParse(wb.getCellValue(i, 16), out line15);
                mine.ApprovedDimOperationalMeasureKey = line15;

                records.Add(mine);
                
            }
            return records;
        }
        
    }
}
