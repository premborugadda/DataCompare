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
using System.Data;

namespace DataCompare
{
    public class Program
    {

        public static void Main(string[] args)
        {
            string homedir, hanaSheetLocation, approvalSheetLocation, 
                naidSheetLocation, resultSheet, budgetSheetLoctation;            
            char[] delimiter = { '\t' };
            List<Result> resMines = new List<Result>();
            DataCompare.Report.Report report = new Report.Report();
            

            homedir = "C:\\Users\\pborugadda\\Documents\\Vale\\";
            hanaSheetLocation = homedir + "KPI_EXTRACT_FULL_V4 24th Nov 2020.txt";
            approvalSheetLocation = homedir + "Approval_Extract_2020_Jan_Nov.xlsx";
            naidSheetLocation = homedir + "2020.Nov.18 NA Integrated Dashboard.xlsm";
            budgetSheetLoctation = homedir + "2020 BM Production Budget - R4V2.xlsx";
            resultSheet = homedir + "Daily Production Dashboard Validation.xlsm";

            Functions.Helpers.KillExcel();                      

            //##################################################################################################
            //Reading Approval extract sheet

            DataParser dataParser1 = new DataParser();
            List <Approval> approvalData = new List<Approval>();            
            approvalData = dataParser1.ReadValues(approvalSheetLocation);

            string[] mineNamesApproval = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "OVOID MINE", "THOMPSON MINE", "TOTTEN MINE" };
            Double[] dimOMKeys = { 14046, 14048, 14047, 14049, 77580, 77428, 14050 };
            string fromDateApp = "2020-11-01";
            string toDateApp = "2020-11-16";
            string yearStartDateApp = "2020-01-01";
            int i = 0;
            Console.WriteLine("**** Reading Approval Data ****");
            
            foreach (string mineNameApp in mineNamesApproval)
            {
                Result resultData = new Result(mineNameApp);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                resultData.ResMineName = mineNameApp;
                resultData.Actual.AppDay = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDateApp), Convert.ToDateTime(toDateApp)); ;
                resultData.Actual.AppMon = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDateApp), Convert.ToDateTime(toDateApp)); 
                resultData.Actual.AppYear = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDateApp), Convert.ToDateTime(toDateApp));
                resMines.Add(resultData);
                i++;                
            }
            //##################################################################################################
            //Reading NAID Dashboard sheet

            NAID naid = new NAID(naidSheetLocation);

            naid.createMine("COLEMAN MINE", 14);
            naid.createMine("COPPER CLIFF MINE", 16);
            naid.createMine("CREIGHTON MINE", 18);
            naid.createMine("GARSON MINE", 20);
            naid.createMine("TOTTEN MINE", 22);
            naid.createMine("MANITOBA MINE", 24);
            Console.WriteLine("**** NAID Data ****");            
            string[] minenamesNAID = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "TOTTEN MINE", "MANITOBA MINE" };
            
            
            foreach (string mineNaid in minenamesNAID)
            {
                Mine curMine = naid.getMine(mineNaid);
                Result resultData = report.getResMine(mineNaid);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                resultData.Actual.NaidDay = Convert.ToDouble(curMine.Production.oreDayA);
                resultData.Actual.NaidMon = Convert.ToDouble(curMine.Production.oreMTDA);
                resultData.Actual.NaidYear = Convert.ToDouble(curMine.Production.oreYTDA);

                resultData.Budget.NaidDay = Convert.ToDouble(curMine.Production.oreDayB);
                resultData.Budget.NaidMon = Convert.ToDouble(curMine.Production.oreMTDB);
                resultData.Budget.NaidYear = Convert.ToDouble(curMine.Production.oreYTDB);

                resMines.Add(resultData);
                
            }

            

            //##################################################################################################
            //Reading HANA extract sheet

            string fileContent = Helpers.ReadFileContent(hanaSheetLocation);
            List<Hana> data = new List<Hana>();
            Parser dataParser = new Parser();
            data = dataParser.GetData(fileContent);

            string[] minenamesHANA = { "COLEMAN MINE", "COPPER CLIFF MINE", "CREIGHTON MINE", "GARSON MINE", "OVOID MINE", "THOMPSON MINE", "TOTTEN MINE" };
            string fromDate = "2020-11-01";
            string toDate = "2020-11-16";
            string yearStartDate = "2020-01-01";
            Console.WriteLine("**** HANA Data ****");
            
            foreach (string mineHana in minenamesHANA)
            {
                Result resultData = report.getResMine(mineHana);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                resultData.Actual.HanaDay = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaMon = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Actual.HanaYear = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));

                resultData.Budget.HanaDay = Helpers.SumOfValues(data, 1.3, mineHana, "BUDGET", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                resultData.Budget.HanaMon = Helpers.SumOfValues(data, 1.3, mineHana, "BUDGET", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                resultData.Budget.HanaYear = Helpers.SumOfValues(data, 1.3, mineHana, "BUDGET", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));
                resMines.Add(resultData);
            }

            //##################################################################################################
            //Reading Budget sheet

            Console.WriteLine("**** Budget Data ****");
            dataParse dt = new dataParse();
            List<Budget> budgetData = new List<Budget>();
            budgetData = dt.GetValues(budgetSheetLoctation);
            //string[] mineNamesBudget = { "Coleman", "Copper Cliff North", "Creighton", "Garson", "Gertrude", "Ellen", "Stobie", "Totten", "Garson Ramp", "OB & CCM Extra" };
            string[] mineNamesBudget = { "Coleman", "Copper Cliff North", "Creighton", "Garson", "Gertrude", "Ellen", "Totten" };

            
            foreach (string mineNameBud in mineNamesBudget)
            {
                Result resultData = report.getResMine(mineNameBud);
                resultData.Actual = new Actual();
                resultData.Budget = new BudgetValues();
                double BudValue = Helpers.BudgetMineValue(budgetData, mineNameBud, 11);
                string month = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(11);

                //objReport.ReportDouble(resultSheet, 2, iRow, 11, 0, curMine.Production.oreDayA);
                Console.WriteLine(mineNameBud + ": Budget value for the month of " + month + " is " + BudValue);
            }
            Console.WriteLine("\n");




            //##################################################################################################
            //Open the report
            var objReport = new Report.Report();
            int iRow = 3;
            foreach (string mineNameApp in mineNamesApproval)
            {
                objReport.writeData(mineNameApp, resMines);
                
            }
            //close the report
            //objReport.saveCloseFile(resultSheet);
            Functions.Helpers.KillExcel();
            Console.Read();
            
        }

    }
}

