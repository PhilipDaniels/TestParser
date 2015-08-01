using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Parses a trx file (as produced by MS Test) and returns a set of <see cref="TestResult"/> objects.
    /// </summary>
    public class TrxFileParser : ITestFileParser
    {
        /// <summary>
        /// Parses a trx file (as produced by MS Test) and returns a set of <see cref="TestResult"/> objects.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns>Set of test results.</returns>
        /// <exception cref="System.IO.FileNotFoundException">The file ' + filename + ' does not exist.</exception>
        public IEnumerable<TestResult> Parse(string filename)
        {
            filename.ThrowIfFileDoesNotExist("filename");

            XNamespace ns = @"http://microsoft.com/schemas/VisualStudio/TeamTest/2010";
            var doc = XDocument.Load(filename);

            var testDefinitions = (from unitTest in doc.Descendants(ns + "UnitTest")
                         select new
                         {
                            executionId = unitTest.Element(ns + "Execution").Attribute("id").Value,
                            codeBase = unitTest.Element(ns + "TestMethod").Attribute("codeBase").Value,
                            className = unitTest.Element(ns + "TestMethod").Attribute("className").Value,
                            testName = unitTest.Element(ns + "TestMethod").Attribute("name").Value
                         }
                         ).ToList();


            var results = (from utr in doc.Descendants(ns + "UnitTestResult")
                           let executionId = utr.Attribute("executionId").Value
                           let message = utr.Descendants(ns + "Message").FirstOrDefault()
                           let stackTrace = utr.Descendants(ns + "StackTrace").FirstOrDefault()
                           let st = DateTime.Parse(utr.Attribute("startTime").Value).ToUniversalTime()
                           let et = DateTime.Parse(utr.Attribute("endTime").Value).ToUniversalTime()
                           select new TestResult()
                           {
                               TestResultFileType = Core.TestResultFileType.Trx,
                               ResultsPathName = filename,
                               AssemblyPathName = (from td in testDefinitions where td.executionId == executionId select td.codeBase).Single(),
                               FullClassName = (from td in testDefinitions where td.executionId == executionId select td.className).Single(),
                               ComputerName = utr.Attribute("computerName").Value,
                               StartTime = st,
                               EndTime = et,
                               Outcome = utr.Attribute("outcome").Value,
                               TestName = utr.Attribute("testName").Value,
                               ErrorMessage = message == null ? "" : message.Value,
                               StackTrace = stackTrace == null ? "" : stackTrace.Value,
                               DurationInSeconds = (et - st).TotalSeconds
                           }
                           ).OrderBy(r => r.ResultsPathName).
                             ThenBy(r => r.AssemblyPathName).
                             ThenBy(r => r.ClassName).
                             ThenBy(r => r.TestName).
                             ThenBy(r => r.StartTime);

            return results;
        }
    }
}
