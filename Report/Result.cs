using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.Report
{
    class Result
    {
        private Actual actual;
        private BudgetValues budget;
        
        public string ResMineName;
        public Result(string mineName)
        {
            ResMineName = mineName;

            actual = new Actual();
            budget = new BudgetValues();

            //ResMineValues = new Dictionary<string, string>();
        }

        
        public Actual Actual { get; set; }
        public BudgetValues Budget { get; set; }

        

    }
}
