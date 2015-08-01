using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestParser.Core
{
    public class TestResultFactory
    {
        readonly TrxFileParser trxFileParser;
        readonly NUnitFileParser nunitFileParser;

        public TestResultFactory()
        {
            trxFileParser = new TrxFileParser();
            nunitFileParser = new NUnitFileParser();
        }

        public IEnumerable<TestResult> CreateFromTestFiles(IEnumerable<string> testFileNames)
        {
            var results = new List<TestResult>();
            foreach (string file in testFileNames)
            {
                var fileType = TestResultFileTypeGuesser.GuessFileType(file);

                switch (fileType)
                {
                    case TestResultFileType.Trx:
                        results.AddRange(trxFileParser.Parse(file));
                        break;
                    case TestResultFileType.NUnit:
                        break;
                    default:
                        throw new Exception("Unhandled fileType " + fileType);
                }
            }

            return results;
        }
    }
}
