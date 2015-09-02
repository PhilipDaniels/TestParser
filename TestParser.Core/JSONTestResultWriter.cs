using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace TestParser.Core
{
    public class JSONTestResultWriter : ITestResultWriter
    {
        public void WriteResults(Stream s, ParsedData testResults)
        {
            var utf8WithoutBom = new UTF8Encoding(false);

            using (var sw = new StreamWriter(s, utf8WithoutBom))
            {
                var settings = new JsonSerializerSettings();
                settings.Formatting = Formatting.Indented;
                settings.NullValueHandling = NullValueHandling.Ignore;
                settings.Converters.Add(new StringEnumConverter());

                var jser = JsonSerializer.Create(settings);

                jser.Serialize(sw, testResults);
            }
        }
    }
}
