using System;
using System.IO;
using YamlDotNet.Core;
using YamlDotNet.Core.Events;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace DataCompare.ConfigLayer
{
    internal static class YamlConvert
    {
        public static T Deserialize<T>(string input)
        {
            Parser parser;
            try
            {
                parser = new Parser(new StringReader(input));
                parser.Consume<StreamStart>();
            }
            catch (Exception e)
            {
                throw new InvalidOperationException($"Failed to parse input: {e}.");
            }

            var deserializer = new DeserializerBuilder().WithNamingConvention(CamelCaseNamingConvention.Instance).Build();
            T model = default;

            while (parser.TryConsume<DocumentStart>(out var _))
            {
                try
                {
                    model = deserializer.Deserialize<T>(parser);
                }
                catch (Exception e)
                {
                    throw new InvalidOperationException($"Failed to deserialize input: {e}.");
                }
            }

            return model;
        }
    }
}
