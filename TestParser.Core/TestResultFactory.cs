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
        readonly NCrunchCoverageParser nCrunchCoverageParser;

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultFactory"/> class.
        /// </summary>
        public TestResultFactory()
        {
            trxFileParser = new TrxFileParser();
            nunit2FileParser = new NUnit2FileParser();
            nCrunchCoverageParser = new NCrunchCoverageParser();
        }

        /// <summary>
        /// Creates a set of <see cref="TestResult"/> objects from test files.
        /// </summary>
        /// <param name="testFileNames">The test file names.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Unhandled fileType  + fileType</exception>
        public ParsedData CreateResultsFromTestFiles(IEnumerable<string> testFileNames)
        {
            var results = new ParsedData();

            foreach (string file in testFileNames)
            {
                var fileType = InputFileTypeGuesser.GuessFileType(file);

                switch (fileType)
                {
                    case InputFileType.Trx:
                        results.AddRange(trxFileParser.Parse(file));
                        break;
                    case InputFileType.NUnit2:
                        results.AddRange(nunit2FileParser.Parse(file));
                        break;
                    case InputFileType.NCrunchCoverage:
                        results.AddRange(nCrunchCoverageParser.Parse(file));
                        break;
                    default:
                        throw new Exception("Unhandled fileType " + fileType);
                }
            }

            results.Summarise();

            return results;
        }
    }
}
