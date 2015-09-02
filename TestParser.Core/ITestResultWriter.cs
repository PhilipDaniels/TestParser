using System.IO;

namespace TestParser.Core
{
    public interface ITestResultWriter
    {
        void WriteResults(Stream s, ParsedData testResults);
    }
}
