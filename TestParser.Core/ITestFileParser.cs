using System.Collections.Generic;

namespace TestParser.Core
{
    public interface ITestFileParser
    {
        IEnumerable<TestResult> Parse(string filename);
    }
}
