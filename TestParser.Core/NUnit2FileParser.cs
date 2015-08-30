using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Parses a file produced by NUnit 2.x and returns a set of <see cref="TestResult"/> objects.
    /// </summary>
    public class NUnit2FileParser : ITestFileParser
    {
        /// <summary>
        /// Parses a file produced by NUnit 2.x and returns a set of <see cref="TestResult"/> objects.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns>Set of test results.</returns>
        /// <exception cref="System.IO.FileNotFoundException">The file ' + filename + ' does not exist.</exception>
        public IEnumerable<TestResult> Parse(string filename)
        {
            filename.ThrowIfFileDoesNotExist("filename");

            var doc = XDocument.Load(filename);

            try
            {
                var results = (from tr in doc.Descendants("test-results")
                               from tc in doc.Descendants("test-case")
                               let message = tc.Descendants("message").SingleOrDefault()
                               let stackTrace = tc.Descendants("stack-trace").SingleOrDefault()
                               let times = tc.Attributes("time")
                               select new TestResult()
                               {
                                   TestResultFileType = TestResultFileType.NUnit2,
                                   ResultsPathName = filename,
                                   AssemblyPathName = tr.Attribute("name").Value,
                                   ComputerName = tr.Element("environment").Attribute("machine-name").Value,
                                   TestName = GetTestName(tc.Attribute("name").Value),
                                   FullClassName = GetFullClassName(tc.Attribute("name").Value),
                                   Outcome = GetOutcome(tc.Attribute("result").Value),
                                   ErrorMessage = message == null ? "" : message.Value,
                                   StackTrace = stackTrace == null ? "" : stackTrace.Value,
                                   DurationInSeconds = times.Any() ? Convert.ToDouble(times.First().Value) : 0.0
                               }
                               ).OrderBy(r => r.ResultsPathName).
                                 ThenBy(r => r.AssemblyPathName).
                                 ThenBy(r => r.ClassName).
                                 ThenBy(r => r.TestName).ToList();

                return results;
            }
            catch (Exception ex)
            {
                throw new ParseException("Error while parsing NUnit file '" + filename + "'", ex);
            }
        }

        string GetOutcome(string resultAttributeValue)
        {
            if (resultAttributeValue.Equals("Success", StringComparison.InvariantCultureIgnoreCase))
                return KnownOutcomes.Passed;
            else if (resultAttributeValue.Equals("Failure", StringComparison.InvariantCultureIgnoreCase))
                return KnownOutcomes.Failed;
            else
                return resultAttributeValue;
        }

        string GetTestName(string nameAttributeValue)
        {
            string className, testName;
            SplitTestName(nameAttributeValue, out className, out testName);
            return testName;
        }

        string GetFullClassName(string nameAttributeValue)
        {
            string className, testName;
            SplitTestName(nameAttributeValue, out className, out testName);
            return className;
        }

        /// <summary>
        /// Splits the name of the test. Given "BassUtils.Tests.StringBuilderExtensionTests.AppendCSV_IfDelimiterAppearsInTerm_DoublesDelimiter"
        /// className is "BassUtils.Tests.StringBuilderExtensionTests" and testName is "AppendCSV_IfDelimiterAppearsInTerm_DoublesDelimiter".
        /// </summary>
        /// <param name="nameAttributeValue">The name attribute value.</param>
        /// <param name="className">Name of the class.</param>
        /// <param name="testName">Name of the test.</param>
        void SplitTestName(string nameAttributeValue, out string className, out string testName)
        {
            // NUnit can create names like 'BassUtils.Tests.ConvTests.StringToBest_ForVarious_ReturnsExpected("456,789.123",45678)'
            // for data-driven tests.
            string testParams = "";
            int idx = nameAttributeValue.IndexOf('(');
            if (idx != -1)
            {
                testParams = nameAttributeValue.Substring(idx);
                nameAttributeValue = nameAttributeValue.Substring(0, idx);
            }

            idx = nameAttributeValue.LastIndexOf('.');
            if (idx == -1)
            {
                className = nameAttributeValue;
                testName = nameAttributeValue;
            }
            else
            {
                className = nameAttributeValue.Substring(0, idx);
                testName = nameAttributeValue.Substring(idx + 1);
            }

            testName += testParams;
        }
    }
}
