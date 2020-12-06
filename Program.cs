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
using DataCompare.ApprovalParser;
using System.Runtime.InteropServices;

namespace DataCompare
{
    public class Program
    {

        public static void Main(string[] args)
        {
            string homedir, hanaSheetLocation, approvalSheetLocation, naidSheetLocation;            
            char[] delimiter = { '\t' };

            homedir = "C:\\Users\\pborugadda\\Documents\\Vale\\";
            hanaSheetLocation = homedir + "KPI_EXTRACT_FULL_V4 24th Nov 2020.txt";
            approvalSheetLocation = homedir + "Approval_Extract_2020_Jan_Nov.xlsx";
            naidSheetLocation = homedir + "2020.Nov.18 NA Integrated Dashboard.xlsm";



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
            Console.WriteLine("**** Approval Data ****");

            foreach (string mineNameApp in mineNamesApproval)
            {                
                int dayValueApp = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(toDateApp), Convert.ToDateTime(toDateApp));
                Console.WriteLine(toDateApp + ": " + mineNameApp.ToString() + " Day Value: " + dayValueApp);

                int monthToDateValueApp = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(fromDateApp), Convert.ToDateTime(toDateApp));
                Console.WriteLine(fromDateApp + " to " + toDateApp + ": " + mineNameApp.ToString() + " Month to Date Value: " + monthToDateValueApp);

                int yearToDateValueApp = Helpers.ApprovalSumOfValues(approvalData, dimOMKeys[i], Convert.ToDateTime(yearStartDateApp), Convert.ToDateTime(toDateApp));
                Console.WriteLine(yearStartDateApp + " to " + toDateApp + ": " + mineNameApp.ToString() + " Year to Date Value: " + yearToDateValueApp);
                Console.WriteLine("\n");
                i++;
            }

            //##################################################################################################
            //Reading NAID Dashboard sheet

            NAID naid = new NAID(naidSheetLocation);

            naid.createMine("COLEMAN", 14);
            naid.createMine("COPPER CLIFF", 16);
            naid.createMine("CREIGHTON", 18);
            naid.createMine("GARSON", 20);
            naid.createMine("TOTTEN", 22);
            naid.createMine("MANITOBA", 24);
            Console.WriteLine("**** NAID Data ****");

            string[] minenamesNAID = { "COLEMAN", "COPPER CLIFF", "CREIGHTON", "GARSON", "TOTTEN", "MANITOBA" };
            foreach (string mineNaid in minenamesNAID)
            {
                Mine curMine = naid.getMine(mineNaid);

                Console.WriteLine(curMine.MineName + " Day Value: " + curMine.Production.oreDayA);
                Console.WriteLine(curMine.MineName + " Month to Date Value: " + curMine.Production.oreMTDA);
                Console.WriteLine(curMine.MineName + " Year to Date Value: " + curMine.Production.oreYTDA);
                Console.WriteLine("\n");
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
                int dayValue = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(toDate), Convert.ToDateTime(toDate));
                Console.WriteLine(toDate + ": " + mineHana.ToString() + " Day Value: " + dayValue);

                int mtdSum = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(fromDate), Convert.ToDateTime(toDate));
                Console.WriteLine(fromDate + "to " + toDate + ": " + mineHana.ToString() + " Month to Date Value: " + mtdSum);
                
                int ytdSum = Helpers.SumOfValues(data, 1.3, mineHana, "ACTUAL", Convert.ToDateTime(yearStartDate), Convert.ToDateTime(toDate));
                Console.WriteLine(yearStartDate + "to " + toDate + ": " + mineHana.ToString() + " Year to Date Value: " + ytdSum);
                Console.WriteLine("\n");
            }
            
            Console.Read();
            
        }

    }
}

