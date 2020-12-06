using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCompare.NAIDParser
{
    class Mine
    {
        private Risk risk;
        private Safety safety;
        private Env env;
        private People people;
        private Production production;
        private Maintenance maintenance;
        private Costs costs;
        private Finance finance;

        public Mine(string mineName, int indexColumn)
        {
            MineName = mineName;
            IndexColumn = indexColumn;

            risk = new Risk();
            safety = new Safety();
            env = new Env();
            people = new People();
            production = new Production();
            maintenance = new Maintenance();
            costs = new Costs();
            finance = new Finance();

            MineValues = new Dictionary<string, string>();
        }

        public string MineName { get; set; }

        public int IndexColumn { get; set; }

        public Risk Risk { get; set; }
        public Safety Safety { get; set; }
        public Env Env { get; set; }
        public People People { get; set; }
        public Production Production { get; set; }
        public Maintenance Maintenance { get; set; }
        public Costs Costs { get; set; }
        public Finance Finance { get; set; }

        public Dictionary<string, string> MineValues { get; }

        public void addValueToDictionary(string key, string value)
        {
            this.MineValues.Add(key, value);
        }


    }
}
