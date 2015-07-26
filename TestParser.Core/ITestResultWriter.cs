using System.Collections.Generic;
using System.IO;

namespace TestParser.Core
{
    public interface ITestResultWriter
    {
        void WriteResults(Stream s, IEnumerable<TestResult> testResults);
    }
}
