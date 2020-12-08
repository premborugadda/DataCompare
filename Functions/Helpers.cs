using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataCompare.ApprovalParser;
using DataCompare.BudgetParser;
using DataCompare.HanaParser;

namespace DataCompare.Functions
{
    public static class Helpers
    {
        public static int SumOfValues(this List<Hana> data, double kpiID, string location, string kpiTYPE, DateTime fromDate, DateTime toDate)
        {
            var test = data.Where(x =>  x.KPI_ID.Equals(kpiID) && x.KPI_LOCATION.Equals(location) && x.KPI_TYPE.Equals(kpiTYPE) && x.CTRL_DATE_KEY >= fromDate && x.CTRL_DATE_KEY <= toDate).ToList();
            return test.Select(y => Convert.ToInt32(y.KPI_NON_RATIO_VALUES)).Sum(); ;
        }
        public static string ReadFileContent(string filePath)
        {
            StreamReader file = new StreamReader(filePath);
            string content = file.ReadToEnd();
            file.Dispose();
            file.Close();
            return content;
        }

        public static int ApprovalSumOfValues(this List<Approval> data, double dimOMKey, DateTime fromDate, DateTime toDate)
        {
            var test1 = data.Where(x => x.DimOperationalMeasureKey.Equals(dimOMKey) && x.CalendarDate >= fromDate && x.CalendarDate<= toDate).ToList();            
            return test1.Select(y => Convert.ToInt32(y.RevisedVoulmeDry)).Sum(); ;
        }

        public static Double BudgetMineValue(this List<Budget> data, string mineName, int month)
        {            
            string monthName = DateTimeFormatInfo.CurrentInfo.GetMonthName(month);
            var test1 = data.Where(x => x.mineName.Equals(mineName) && x.budgetMonth.ToLower().Equals(monthName.ToLower())).ToList();
            return test1.Select(y => Convert.ToDouble(y.budgetValue)).First();
        }
        public static void KillExcel()
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    clsProcess.Kill();                    
                    //break;                    
                }
            }
        }
    }
}
