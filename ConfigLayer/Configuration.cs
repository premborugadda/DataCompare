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

        public List<Dates> Dates { get; set; }
        public List<Mines> Mines { get; set; }
        public List<DimOMKeys> DimOMKeys { get; set; }
    }

    public class Dates
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class Mines
    {
        public string Mine { get; set; }
    }

    public class DimOMKeys
    {
        public string Key { get; set; }
    }
}
