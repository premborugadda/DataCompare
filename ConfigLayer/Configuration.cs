using System.IO;
using System.Collections.Generic;
using System.Reflection;

namespace DataCompare.ConfigLayer
{
    public class Configuration
    {
        public static Configuration Load()
        {
            var path = Config();

            var o = YamlConvert.Deserialize<Configuration>(File.ReadAllText(path));

            return o;

            string Config() => Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".config.yml");
        }
        
        public List<Paths> Paths { get; set; }
        public List<Dates> Dates { get; set; }
        public List<Kpiloc> Kpiloc { get; set; }
        public List<KpiID> KpiID { get; set; }
        public List<DimOMKeys> DimOMKeys { get; set; }
    }

    public class Paths
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }
    public class Dates
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class Kpiloc
    {
        public string Plant { get; set; }
    }

    public class KpiID
    {
        public double Id { get; set; }
    }

    public class DimOMKeys
    {
        public double Keynum { get; set; }
    }
}
