using DataCompare.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.NAIDParser
{
    class NAID
    {
        private string sheetLocation;
        private Mine m;
        private List<Mine> mines;
        private ExcelWorkbook wb;

        public NAID(string sheetLocation)
        {
            this.sheetLocation = sheetLocation;
            this.mines = new List<Mine>();

            this.wb = new ExcelWorkbook(sheetLocation);
        }

        public void createMine(string mineName, int indexColumn)
        {
            try
            {
                m = new Mine(mineName, indexColumn);

                readMineValues(m);
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        // TODO fill readMineValues()
        public void readMineValues(Mine mine)
        {
            mine.Risk = new Risk();
            mine.Production = new Production();
            // TODO open sheet
            //wb.openWorkbook();
            //wb.openSheet(1);
            var xlRange = wb.xlRange;
            var rowCount = wb.xlRange.Rows.Count;


            for (int i = 1; i < rowCount; i++)
            {
                if (xlRange.Cells[i, m.IndexColumn] != null && xlRange.Cells[i, m.IndexColumn].Value2 != null)
                {
                    

                    //Risk data
                    if (i == 6)
                    {
                        m.Risk.NumHIRAsImmediateValueA = wb.getCellValue(i, m.IndexColumn);
                        m.Risk.NumHIRAsImmediateValueB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 7)
                    {
                        // m.Risk.NumHIRADateValue = xlRange.Cells[i, m.IndexColumn].value2;
                        m.Risk.NumHIRADateValue = wb.getCellValue(i, m.IndexColumn);

                    }

                    //Production data
                    if (i == 35)
                    {
                        m.Production.oreDayA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.oreDayB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 36)
                    {
                        m.Production.oreMTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.oreMTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 37)
                    {
                        m.Production.oreYTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.oreYTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    // Nickel values
                    if (i == 41)
                    {
                        m.Production.nickDayA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.nickDayB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 42)
                    {
                        m.Production.nickMTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.nickMTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 43)
                    {
                        m.Production.nickYTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.nickYTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    //Copper values
                    if (i == 47)
                    {
                        m.Production.copperDayA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.copperDayB= wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 48)
                    {
                        m.Production.copperMTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.copperMTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }

                    if (i == 49)
                    {
                        m.Production.copperYTDA = wb.getCellValue(i, m.IndexColumn);
                        m.Production.copperYTDB = wb.getCellValue(i, m.IndexColumn + 1);
                    }
                }
            }


            
            mines.Add(m);
        }

        public Mine getMine(string mineName)
        {
            return mines.Find(mine => mine.MineName == mineName);
        }

        public List<Mine> Mines { get; }

        // TODO Define where measure rows are in NAID

        // Meta
        private static int updateCol = 5;

        private static int mineName = 3;
        private static int manager = 4;

        // RISK
        private static int riskNumHiras = 6;
        private static int riskHiraDate = 7;

        // SAFETY
        private static int safetyNumOccurDay = 0;
        private static int safetyNumOccurMTD = 0;
        private static int safetyNumOccurYTP = 0;

        private static int safetyNumIncidentsDay = 0;
        private static int safetyNumIncidentsMTD = 0;
        private static int safetyNumIncidentsYTD = 0;

        private static int safetyNumRepInjuriesM1 = 0;
        private static int safetyNumRepInjuriesYTD = 0;

        private static int safetyNumRecInjuriesM1 = 0;
        private static int safetyNumRecInjuriesYTD = 0;

        private static int safetyN2sM1 = 0;
        private static int safetyN2YTD = 0;

        private static int safteyYTDTRDR = 0;
        private static int safteyYTDTRIFR = 0;
        private static int safteyYTDAIFR = 0;

        // ENV
        private static int envCriticalIncidents = 0;
        private static int envHPIEnvironment = 0;
        private static int envLegalNotifNonConf = 0;
        private static int envGHG = 0;

        // PEOPLE
        private static int peopleEmployees = 0;
        private static int peopleContractors = 0;
        private static int peopleAbsenteeism = 0;
        private static int peopleDnINumFemale = 0;

        // PRODUCTION
        private static int productionHoistedOreMilledDay = 0;
        private static int productionHoistedOreMilledMTD = 0;
        private static int productionHoistedOreMilledYTD = 0;
        private static int productionHoistedOreMilledFYPlan = 0;
        private static int productionHoistedOreMilledFYBudget = 0;

        private static int productionContainedDay = 0;
        private static int productionContainedMTD = 0;
        private static int productionContainedYTD = 0;
        private static int productionContainedFYPlan = 0;
        private static int productionContainedFYBudget = 0;

        // MAINTENANCE
        private static int maintenanceMPA = 0;
        private static int maintenanceSMA = 0;
        private static int maintenanceCMA = 0;
        private static int maintenanceLabourUtilization = 0;
        private static int maintenanceAvailability = 0;

        // COSTS
        private static int costsPrimaryMTH = 0;
        private static int costsPrimaryYTD = 0;
        private static int costsPrimaryFYPlan = 0;
        private static int costsPrimaryFYBudget = 0;

        private static int costsFixedCostMTH = 0;
        private static int costsFixedCostYTD = 0;
        private static int costsFixedCostFYPlan = 0;
        private static int costsFixedCostFYBudget = 0;

        private static int costsTotalMTH = 0;
        private static int costsTotalYTD = 0;
        private static int costsTotalFYPlan = 0;
        private static int costsTotalFYBudget = 0;

        // FINANCE
        private static int financeSusCapexAccr = 73;
        private static int financeSusCapexCash = 74;
        private static int financeEBITDA = 0;
        private static int financeEBITDAAdj = 0;
        private static int financeTNiYTD = 0;
        private static int financeTNiLessByProd = 0;
        private static int financeTYTDHoistedProc = 0;
        private static int financeUnitCashCost = 80;
    }
}
