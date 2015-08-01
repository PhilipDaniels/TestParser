using System;
using System.Collections.Generic;

namespace TestParser.Core
{
    /// <summary>
    /// Passes each test file to an appropriate parser and accumulates a list
    /// of all the <see cref="TestResult"/> objects it can gather.
    /// </summary>
    public class TestResultFactory
    {
        readonly TrxFileParser trxFileParser;
        readonly NUnit2FileParser nunit2FileParser;

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultFactory"/> class.
        /// </summary>
        public TestResultFactory()
        {
            trxFileParser = new TrxFileParser();
            nunit2FileParser = new NUnit2FileParser();
        }

        /// <summary>
        /// Creates a set of <see cref="TestResult"/> objects from test files.
        /// </summary>
        /// <param name="testFileNames">The test file names.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Unhandled fileType  + fileType</exception>
        public IEnumerable<TestResult> CreateResultsFromTestFiles(IEnumerable<string> testFileNames)
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
                    case TestResultFileType.NUnit2:
                        results.AddRange(nunit2FileParser.Parse(file));
                        break;
                    default:
                        throw new Exception("Unhandled fileType " + fileType);
                }
            }

            return results;
        }
    }
}
