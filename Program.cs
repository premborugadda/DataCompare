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
using DataCompare.ConfigLayer;

namespace DataCompare
{
    public class Program
    {

        public static void Main(string[] args)
        {
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Program execution started ***** ");


            //##################################################################################################
            //YAML Config
           
            //##################################################################################################
            //Closing Excel instances           

            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Closing Excel instances ***** ");
            Functions.Helpers.KillExcel();

            //##################################################################################################
            //YAML Config
            var config = Configuration.Load();
            string homedir, hanaSheetLocation, approvalSheetLocation,
                naidSheetLocation, resultSheet, budgetSheetLoctation;

            var kpiLocs = config.Kpiloc;
            var KPI_IDs = config.KpiID;
            var dimKeys = config.DimOMKeys;
            
            List<string> mineNames = new List<string>();
            List<double> dimOMKeys = new List<double>();

            foreach (var dimKey in dimKeys)
            {
                dimOMKeys.Add(Convert.ToDouble(dimKey.Keynum));                
            }

            foreach (var kpilocation in kpiLocs)
            {
                mineNames.Add(kpilocation.Plant.ToString());                
            }

            //string[] mineNames = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "OVOID MINE", "THOMPSON MINE", "TOTTEN MINE" };
            //Double[] dimOMKeys = { 14046, 14048, 14047, 14049, 77580, 77428, 14050 };

            var yearStartDate = config.Dates?.FirstOrDefault(p => p.Key == "yearStartDate")?.Value;
            var fromDate = config.Dates?.FirstOrDefault(p => p.Key == "from")?.Value;
            DateTime prevMonthEndDate = Convert.ToDateTime(fromDate).AddDays(-1);
            var toDate = config.Dates?.FirstOrDefault(p => p.Key == "to")?.Value;
            var testhomedir = config.Paths?.FirstOrDefault(p => p.Key == "homedir")?.Value;
            homedir = testhomedir + toDate + "\\";
            var objFolder = new ReadFileNames();
            string[] fileNames = objFolder.getSourceFileNames(homedir);

            hanaSheetLocation = fileNames[0]; 
            approvalSheetLocation = fileNames[1]; 
            naidSheetLocation = fileNames[2]; 
            budgetSheetLoctation = fileNames[3]; 
            resultSheet = fileNames[4];
            
            var reportDate = Convert.ToDateTime(System.DateTime.ParseExact(toDate, "yyyy-MM-dd", new CultureInfo("en-us")));
            var daysInMonth = DateTime.DaysInMonth(reportDate.Year, reportDate.Month);
            var curDate = reportDate.Day;

            //##################################################################################################
            //Initializing variables

            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading input parameters ***** ");
            List<Result> resMines = new List<Result>();
            DataCompare.Report.Report report = new Report.Report();

            //##################################################################################################
            //Reading Approval extract sheet

            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading Approval Extract data ***** ");
            DataParser dataParser1 = new DataParser();
            List <Approval> approvalData = new List<Approval>();            
            approvalData = dataParser1.ReadValues(approvalSheetLocation);
            
            //##################################################################################################
            //Reading NAID Dashboard sheet
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading NAID Dashboard data ***** ");
            int iCounter = 0;
            NAID naid = new NAID(naidSheetLocation);
            int[] intCol = { 14, 16, 18, 20, 26, 24, 22 };
            foreach (string mineName in mineNames)
            {
                naid.createMine(mineName, intCol[iCounter]);
                iCounter++;
            }
           
            //##################################################################################################
            //Reading Budget sheet

            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading Budget data ***** ");

            dataParse dt = new dataParse();
            List<Budget> budgetData = new List<Budget>();
            budgetData = dt.GetValues(budgetSheetLoctation, mineNames);
            
                       
            //##################################################################################################
            //Reading HANA extract sheet

            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading HANA Full Extract data ***** ");
            string fileContent = Helpers.ReadFileContent(hanaSheetLocation);
            List<Hana> data = new List<Hana>();
            Parser dataParser = new Parser();
            data = dataParser.GetData(fileContent);

            // ################ Reading KPI = 1.3 - Ore Mined volume
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Reading Ore Mined Volume ***** ");
            int i = 0;
            foreach (string mineName in mineNames)
            {
                Result resultData = new Result(mineName);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                Mine curMine = naid.getMine(mineName);
                resultData.ResMineName = mineName;                

                resultData.Actual.AppDay = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDate), Convert.ToDateTime(toDate)); ;
                resultData.Actual.AppMon = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Actual.AppYear = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

                resultData.Budget.AppDay = Math.Round(Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth, 1);
                resultData.Budget.AppMon = Math.Round((Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth) * curDate, 1);
                resultData.Budget.AppYear = Math.Round(Helpers.BudgetMineYTD(budgetData, mineName, reportDate.Month), 1);

                resultData.Actual.HanaDay = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaMon = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaYear = Helpers.SumOfValues(data, 1.3, mineName, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

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
            //Open the report and write KPI 1.3 values
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Writing KPI ID 1.3 data to Output Report ***** ");
            var objReport = new Report.Report();            
            foreach (string mineName in mineNames)
            {
                objReport.writeData(resultSheet, mineName, resMines, toDate, 4);                
            }

            //##################################################################################################
            //Open the report and write KPI 1.1 values
            List<Result> resMines1 = new List<Result>();
            i = 0;
            foreach (string mineName in mineNames)
            {
                Result resultData1 = new Result(mineName);
                resultData1.Actual = new Actual();
                resultData1.Budget = new BudgetValues();
                Mine curMine = naid.getMine(mineName);
                resultData1.ResMineName = mineName;

                //resultData1.Actual.AppDay = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDate), Convert.ToDateTime(toDate)); ;
                //resultData1.Actual.AppMon = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                //resultData1.Actual.AppYear = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

                //resultData1.Budget.AppDay = Math.Round(Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth, 1);
                //resultData1.Budget.AppMon = Math.Round((Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth) * curDate, 1);
                //resultData1.Budget.AppYear = Math.Round(Helpers.BudgetMineYTD(budgetData, mineName, reportDate.Month), 1);

                resultData1.Actual.HanaDay = Helpers.SumOfValues(data, 1.1, mineName, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData1.Actual.HanaMon = Helpers.SumOfValues(data, 1.1, mineName, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData1.Actual.HanaYear = Helpers.SumOfValues(data, 1.1, mineName, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

                resultData1.Budget.HanaDay = Helpers.SumOfValues(data, 1.1, mineName, "BUDGET", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData1.Budget.HanaMon = Helpers.SumOfValues(data, 1.1, mineName, "BUDGET", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData1.Budget.HanaYear = Helpers.SumOfValues(data, 1.1, mineName, "BUDGET", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData1.Actual.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.nickDayA), 1);
                resultData1.Actual.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.nickMTDA), 1);
                resultData1.Actual.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.nickYTDA), 1);

                resultData1.Budget.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.nickDayB), 1);
                resultData1.Budget.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.nickMTDB), 1);
                resultData1.Budget.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.nickYTDB), 1);

                resMines1.Add(resultData1);
                i++;
            }

            //##################################################################################################
            //Open the report and write KPI 1.1 values
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Writing KPI ID 1.1 data to Output Report ***** ");
            //var objReport = new Report.Report();
            foreach (string mineName in mineNames)
            {
                objReport.writeData(resultSheet, mineName, resMines1, toDate, 11);
            }


            List<Result> resMines2 = new List<Result>();
            i = 0;
            foreach (string mineName in mineNames)
            {
                Result resultData2 = new Result(mineName);
                resultData2.Actual = new Actual();
                resultData2.Budget = new BudgetValues();
                Mine curMine = naid.getMine(mineName);
                resultData2.ResMineName = mineName;

                //resultData2.Actual.AppDay = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDate), Convert.ToDateTime(toDate)); ;
                //resultData2.Actual.AppMon = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                //resultData2.Actual.AppYear = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

                //resultData2.Budget.AppDay = Math.Round(Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth, 1);
                //resultData2.Budget.AppMon = Math.Round((Helpers.BudgetMineValue(budgetData, mineName, reportDate.Month) / daysInMonth) * curDate, 1);
                //resultData2.Budget.AppYear = Math.Round(Helpers.BudgetMineYTD(budgetData, mineName, reportDate.Month), 1);

                resultData2.Actual.HanaDay = Helpers.SumOfValues(data, 1.2, mineName, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData2.Actual.HanaMon = Helpers.SumOfValues(data, 1.2, mineName, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData2.Actual.HanaYear = Helpers.SumOfValues(data, 1.2, mineName, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(prevMonthEndDate));

                resultData2.Budget.HanaDay = Helpers.SumOfValues(data, 1.2, mineName, "BUDGET", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData2.Budget.HanaMon = Helpers.SumOfValues(data, 1.2, mineName, "BUDGET", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData2.Budget.HanaYear = Helpers.SumOfValues(data, 1.2, mineName, "BUDGET", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData2.Actual.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.copperDayA), 1);
                resultData2.Actual.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.copperMTDA), 1);
                resultData2.Actual.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.copperYTDA), 1);

                resultData2.Budget.NaidDay = Math.Round(Convert.ToDouble(curMine.Production.copperDayB), 1);
                resultData2.Budget.NaidMon = Math.Round(Convert.ToDouble(curMine.Production.copperMTDB), 1);
                resultData2.Budget.NaidYear = Math.Round(Convert.ToDouble(curMine.Production.copperYTDB), 1);

                resMines2.Add(resultData2);
                i++;
            }

            //##################################################################################################
            //Open the report and write KPI 1.2 values
            Console.WriteLine(" ***** " + DateTime.Now.ToString("h:mm tt") + " ***** Writing KPI ID 1.2 data to Output Report ***** ");
            //var objReport = new Report.Report();
            foreach (string mineName in mineNames)
            {
                objReport.writeData(resultSheet, mineName, resMines2, toDate, 18);
            }

            //close the report
            //objReport.saveCloseFile(resultSheet);
            Functions.Helpers.KillExcel();
            //Console.Read();
            
        }

    }
}

