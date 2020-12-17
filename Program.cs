using DataCompare.ApprovalParser;
using DataCompare.Functions;
using DataCompare.HanaParser;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataCompare.NAIDParser;
using System.Runtime.InteropServices;
using DataCompare.Report;
using DataCompare.BudgetParser;
using DataCompare.ConfigLayer;
using System.Data;

namespace DataCompare
{
    public class Program
    {

        public static void Main(string[] args)
        {


            //##################################################################################################
            //YAML Config
            var config = Configuration.Load();

            //var fromDate = config.Dates?.FirstOrDefault(p => p.Key == "from")?.Value;
            //var toDate = config.Dates?.FirstOrDefault(p => p.Key == "to")?.Value;
            //var yearStartDate = config.Dates?.FirstOrDefault(p => p.Key == "yearStartDate")?.Value;

            //var mines = config.Mines;

            //var dimOMKeys = config.DimOMKeys;
            //##################################################################################################









            //##################################################################################################
            //Initializing variables

            Console.WriteLine(" ***** Reading input parameters ***** ");
            string homedir, hanaSheetLocation, approvalSheetLocation,
                naidSheetLocation, resultSheet, budgetSheetLoctation;
            char[] delimiter = { '\t' };
            List<Result> resMines = new List<Result>();
            DataCompare.Report.Report report = new Report.Report();
            
            //homedir = "C:\\ValeDataTesting\\";
            //var MyIni = new IniFile(homedir + "config.ini");

            //var mines = MyIni.IniReadValue("mineNames","env");
            //var date = MyIni.IniReadValue("reportRunDate","env");


            string[] mineNames = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "TOTTEN MINE", "THOMPSON MINE", "OVOID MINE" };
            Double[] dimOMKeys = { 14046, 14048, 14047, 14049, 14050, 77428, 77580 };
            string fromDate = "2020-11-01";
            string toDate = "2020-11-17";
            string yearStartDate = "2020-01-01";
            int i = 0;
            var reportDate = Convert.ToDateTime(System.DateTime.ParseExact(toDate, "yyyy-MM-dd", new CultureInfo("en-us")));
            var daysInMonth = DateTime.DaysInMonth(reportDate.Year, reportDate.Month);
            var curDate = reportDate.Day;

            //##################################################################################################
            //Reading input parameters into variables

            homedir = "C:\\Users\\pborugadda\\Documents\\Vale\\Nov17\\";
            hanaSheetLocation = homedir + "KPI_EXTRACT_FULL_V4 24th Nov 2020.txt";
            approvalSheetLocation = homedir + "Approval_Extract_2020_Jan_Nov.xlsx";
            naidSheetLocation = homedir + "2020.Nov.18 NA Integrated Dashboard.xlsm";
            budgetSheetLoctation = homedir + "2020 BM Production Budget - R4V2.xlsx";
            resultSheet = homedir + "Daily Production Dashboard Validation.xlsm";

            Console.WriteLine(" ***** Closing Excel instances ***** ");
            Functions.Helpers.KillExcel();
            

            //##################################################################################################
            //Reading Approval extract sheet

            Console.WriteLine(" ***** Reading Approval Extract data ***** ");
            DataParser dataParser1 = new DataParser();
            List <Approval> approvalData = new List<Approval>();            
            approvalData = dataParser1.ReadValues(approvalSheetLocation);
            

            //##################################################################################################
            //Reading NAID Dashboard sheet
            Console.WriteLine(" ***** Reading NAID Dashboard data ***** ");

            NAID naid = new NAID(naidSheetLocation);
            int intCol = 14;
            foreach (string mineName in mineNames)
            {
                naid.createMine(mineName, intCol);
                intCol = intCol + 2;
            }
            //naid.createMine("COPPER CLIFF MINE", 16);
            //naid.createMine("CREIGHTON MINE", 18);
            //naid.createMine("GARSON MINE", 20);
            //naid.createMine("OVOID MINE", 26);
            //naid.createMine("THOMPSON MINE", 24);
            //naid.createMine("TOTTEN MINE", 22);
            //##################################################################################################
            //Reading Budget sheet

            Console.WriteLine(" ***** Reading Budget data ***** ");

            dataParse dt = new dataParse();
            List<Budget> budgetData = new List<Budget>();
            budgetData = dt.GetValues(budgetSheetLoctation, mineNames);
            
            //foreach (string mineNameBud in mineNames)
            //{
            //    Result resultData = new Result(mineNameBud);
            //    resultData.Actual = new Actual();
            //    resultData.Budget = new BudgetValues();
            //    resultData = report.getResMine(mineNameBud);

            //    string month = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(11);
            //    Console.WriteLine(mineNameBud + ": Budget value for the month of " + month + " is " + BudValue);
            //}
            
            //##################################################################################################
            //Reading HANA extract sheet

            Console.WriteLine(" ***** Reading HANA Full Extract data ***** ");
            string fileContent = Helpers.ReadFileContent(hanaSheetLocation);
            List<Hana> data = new List<Hana>();
            Parser dataParser = new Parser();
            data = dataParser.GetData(fileContent);

            // ################ Reading KPI = 1.3 - Ore Mined volume
            Console.WriteLine(" ***** Reading Ore Mined Volume ***** ");

            foreach (string mineName in mineNames)
            {
                Result resultData = new Result(mineName);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                Mine curMine = naid.getMine(mineName);
                resultData.ResMineName = mineName;                

                resultData.Actual.AppDay = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDate), Convert.ToDateTime(toDate)); ;
                resultData.Actual.AppMon = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Actual.AppYear = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData.Budget.AppDay = Math.Round(Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth, 1);
                resultData.Budget.AppMon = Math.Round((Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth) * curDate, 1);
                resultData.Budget.AppYear = Math.Round(Helpers.BudgetMineYTD(budgetData, mineName, reportDate.Month), 1);

                resultData.Actual.HanaDay = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaMon = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaYear = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData.Budget.HanaDay = Helpers.SumOfValues(data, 1.3, mineName, "BUDGET", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData.Budget.HanaMon = Helpers.SumOfValues(data, 1.3, mineName, "BUDGET", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Budget.HanaYear = Helpers.SumOfValues(data, 1.3, mineName, "BUDGET", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData.Actual.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.oreDayA),1);
                resultData.Actual.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.oreMTDA), 1);
                resultData.Actual.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.oreYTDA), 1);

                resultData.Budget.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.oreDayB), 1);
                resultData.Budget.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.oreMTDB), 1);
                resultData.Budget.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.oreYTDB), 1);

                // ################ Reading KPI = 1.1 - Nickel volume
                //Console.WriteLine(" ***** Reading Nickel Volume ***** ");

                resMines.Add(resultData);
                i++;
            }



            //##################################################################################################
            //Open the report
            Console.WriteLine(" ***** Writing data to Output Report ***** ");
            var objReport = new Report.Report();            
            foreach (string mineName in mineNames)
            {
                objReport.writeData(resultSheet, mineName, resMines, toDate);
                
            }
            

            //close the report
            //objReport.saveCloseFile(resultSheet);
            Functions.Helpers.KillExcel();
            //Console.Read();
            
        }

    }
}

