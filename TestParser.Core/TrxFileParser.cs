using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace TestParser.Core
{
    /// <summary>
    /// Parses a trx file (as produced by MS Test) and returns a set of <see cref="TestResult"/> objects.
    /// </summary>
    public class TrxFileParser
    {
        /// <summary>
        /// Parses a trx file (as produced by MS Test) and returns a set of <see cref="TestResult"/> objects.
        /// </summary>
        /// <param name="trxFilename">The TRX filename.</param>
        /// <returns>Set of test results.</returns>
        /// <exception cref="System.IO.FileNotFoundException">The file ' + trxFilename + ' does not exist.</exception>
        public IEnumerable<TestResult> Parse(string trxFilename)
        {
            if (!File.Exists(trxFilename))
                throw new FileNotFoundException("The file '" + trxFilename + "' does not exist.", trxFilename);

            XNamespace ns = @"http://microsoft.com/schemas/VisualStudio/TeamTest/2010";
            var doc = XDocument.Load(trxFilename);

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
                           select new TestResult()
                           {
                               ResultsPathName = trxFilename,
                               AssemblyPathName = (from td in testDefinitions where td.executionId == executionId select td.codeBase).Single(),
                               FullClassName = (from td in testDefinitions where td.executionId == executionId select td.className).Single(),
                               ComputerName = utr.Attribute("computerName").Value,
                               StartTime = DateTime.Parse(utr.Attribute("startTime").Value).ToUniversalTime(),
                               EndTime = DateTime.Parse(utr.Attribute("endTime").Value).ToUniversalTime(),
                               Outcome = utr.Attribute("outcome").Value,
                               TestName = utr.Attribute("testName").Value,
                               ErrorMessage = message == null ? "" : message.Value,
                               StackTrace = stackTrace == null ? "" : stackTrace.Value
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
