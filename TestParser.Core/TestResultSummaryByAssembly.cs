using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace TestParser.Core
{
    [DebuggerDisplay("{AssemblyFileName}, Passed={TotalPassed}/{TotalTests} in {TotalDurationInSeconds} secs.")]
    public class TestResultSummaryByAssembly
    {
        public TestResultSummaryByAssembly()
        {
            Outcomes = new List<ResultOutcome>();
        }

        public string AssemblyPathName { get; set; }

        public string AssemblyFileName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(AssemblyPathName))
                    return "";
                else
                    return Path.GetFileName(AssemblyPathName);
            }
        }

        public List<ResultOutcome> Outcomes { get; private set; }

        public int TotalTests
        {
            get
            {
                return Outcomes.Sum(c => c.NumTests);
            }
        }

        public int TotalPassed
        {
            get
            {
                return Outcomes.Where(c => c.Outcome == ResultOutcome.PassedOutcome).Single().NumTests;
            }
        }

        public double TotalDurationInSeconds
        {
            get
            {
                return Outcomes.Sum(c => c.DurationInSeconds);
            }
        }

        public static IEnumerable<TestResultSummaryByAssembly> Summarise(IEnumerable<TestResult> testResults)
        {
            var summaries = new List<TestResultSummaryByAssembly>();
            var outcomeNames = ResultOutcome.GetOutcomeNames(testResults);

            var sbaRows = testResults.GroupBy(r => r.AssemblyPathName)
                          .Select(gr => new
                          {
                              AssemblyPathName = gr.Key,
                              TestResults = gr
                          })
                          .OrderBy(a => a.AssemblyPathName);

            foreach (var sbaRow in sbaRows)
            {
                var summary = new TestResultSummaryByAssembly();
                summary.AssemblyPathName = sbaRow.AssemblyPathName;

                foreach (var ocn in outcomeNames)
                {
                    var oc = new ResultOutcome() { Outcome = ocn };
                    oc.NumTests = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Count();
                    oc.DurationInSeconds = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Sum(r => r.DurationInSeconds);
                    summary.Outcomes.Add(oc);
                }
                summaries.Add(summary);
            }

            return summaries;
        }
    }
}
