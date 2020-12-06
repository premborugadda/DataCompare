using DataCompare.Functions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataCompare.HanaParser
{
    public class Parser
    {
        public List<Hana> GetData(string csvContent)
        {
            List<Hana> records = new List<Hana>();
            CsvHelper csv = new CsvHelper(csvContent);
            int recordCount = csv.Count();
            bool isFirstRow = true;
            //foreach (string[] line in Regex.Split(csv.ToString(), System.Environment.NewLine).ToList().Where(s => !string.IsNullOrEmpty(s)))

            foreach (string[] line in csv)
            {
                if (isFirstRow)
                {
                    isFirstRow = false;
                }
                else
                {
                    Hana mine = new Hana();

                    double line0 = 0;
                    Double.TryParse(line[0], out line0);
                    mine.KPI_ID = line0;
                    mine.PILLAR = line[1];
                    mine.KPI_CATEGORY = line[2];
                    mine.KPI_DESCRIPTION = line[3];
                    mine.UNIT_OF_MEASURE = line[4];
                    mine.KPI_BETTER_TO = line[5];
                    mine.KPID_ATTRIBUTE_5 = line[6];
                    mine.KPI_LOCATION = line[7];
                    mine.KPI_TYPE = line[8];
                    mine.KPI_TYPE_GROUP = line[9];
                    mine.CALCULATION_METHOD = line[10];
                    mine.KPI_LEVEL_CATEGORY = line[11];
                    mine.RESPONSIBLEO_FOR_COLLECTION = line[12];
                    mine.KPIG_ATTRIBUTE_1 = line[13];
                    mine.KPIG_ATTRIBUTE_3 = line[14];
                    double line15 = 0;
                    Double.TryParse(line[15], out line15);
                    mine.HIERARCHY_ID = line15;
                    mine.HIERARCHY_LEVEL_1 = line[16];
                    mine.HIERARCHY_LEVEL_2 = line[17];
                    mine.HIERARCHY_LEVEL_3 = line[18];
                    mine.HIERARCHY_LEVEL_4 = line[19];
                    mine.HIERARCHY_LEVEL_5 = line[20];
                    mine.HIERARCHY_LEVEL_6 = line[21];
                    mine.HIERARCHY_LEVEL_7 = line[22];
                    mine.HIERARCHY_LEVEL_8 = line[23];
                    mine.HIERARCHY_LATITUDE = line[24];
                    mine.HIERARCHY_LONGITUDE = line[25];
                    mine.CTRL_DATE_KEY = System.DateTime.ParseExact(line[26], "yyyyMMdd", new CultureInfo("en-us"));

                    double line27 = 0;
                    Double.TryParse(line[27], out line27);
                    mine.KPI_NON_RATIO_VALUES = line27;

                    double line28 = 0;
                    Double.TryParse(line[28], out line28);
                    mine.KPI_RATIO_NUMERATOR = line28;

                    double line29 = 0;
                    Double.TryParse(line[29], out line29);
                    mine.KPI_RATIO_DENOMINATOR = line29;

                    double line30 = 0;
                    Double.TryParse(line[30], out line30);
                    mine.KPV1_RAW_KPI_VALUE = line30;

                    double line31 = 0;
                    Double.TryParse(line[31], out line31);
                    mine.KPV2_RAW_KPI_VALUE = line31;

                    //DateTime line32 = null;
                    //DateTime.TryParse(line[32], out line32);

                    //mine.LAST_REFRESHED_DATE = System.DateTime.ParseExact(line[32], "yyyyMMdd", new CultureInfo("en-us"));

                    //char line33 = "";
                    //Char.TryParse(line[33], out line33);
                    mine.XP_DASHBOARD_FLAG = char.Parse(line[33]);

                    double line34 = 0;
                    Double.TryParse(line[34], out line34);
                    mine.ROLL_UP_HIERARCHY_LEVEL = line34;
                    mine.COMMENTS_ACTUAL = line[35];
                    mine.COMMENTS_ACTUALADJ = line[36];
                    records.Add(mine);
                }
                
            }
            return records;
        }

    }
}
