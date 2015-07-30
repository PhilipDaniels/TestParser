using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace TestParser.Core
{
    /// <summary>
    /// A simple class to summarise a set of <see cref="TestResult"/>.
    /// </summary>
    [DebuggerDisplay("{AssemblyFileName}, Passed={TotalPassed}/{TotalTests} in {TotalDurationInSeconds} secs.")]
    public class TestResultSummary : TestResultBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TestResultSummary"/> class.
        /// </summary>
        public TestResultSummary()
        {
            Outcomes = new List<ResultOutcomeSummary>();
        }

        /// <summary>
        /// The outcomes. The well-known outcomes "Passed" and "Failed" will always be
        /// there, and always the first two outcomes. The other outcomes, if any, come
        /// after that.
        /// </summary>
        /// <value>
        /// The outcomes.
        /// </value>
        public List<ResultOutcomeSummary> Outcomes { get; private set; }

        /// <summary>
        /// Gets the total number of tests in this summary line.
        /// </summary>
        /// <value>
        /// The total number of tests in this summary line.
        /// </value>
        public int TotalTests
        {
            get
            {
                return Outcomes.Sum(c => c.NumTests);
            }
        }

        /// <summary>
        /// Gets the total number of passed tests in this summary line.
        /// </summary>
        /// <value>
        /// The total number of passed tests in this summary line.
        /// </value>
        public int TotalPassed
        {
            get
            {
                return Outcomes.Where(c => c.Outcome == ResultOutcomeSummary.PassedOutcome).Single().NumTests;
            }
        }

        /// <summary>
        /// Gets the total duration in seconds in this summary line.
        /// </summary>
        /// <value>
        /// The total duration in seconds in this summary line.
        /// </value>
        public double TotalDurationInSeconds
        {
            get
            {
                return Outcomes.Sum(c => c.TotalDurationInSeconds);
            }
        }

        /// <summary>
        /// Summarises the test results by grouping by assembly.
        /// The <code>ClassName</code> property will be blank for these summary lines.
        /// </summary>
        /// <param name="testResults">The test results.</param>
        /// <returns>Summary of results grouped by assembly.</returns>
        public static IEnumerable<TestResultSummary> SummariseByAssembly(IEnumerable<TestResult> testResults)
        {
            var summaries = new List<TestResultSummary>();
            var outcomeNames = ResultOutcomeSummary.GetOutcomeNames(testResults);

            var sbaRows = testResults.GroupBy(r => r.AssemblyPathName)
                          .Select(gr => new
                          {
                              AssemblyPathName = gr.Key,
                              TestResults = gr
                          })
                          .OrderBy(a => a.AssemblyPathName);

            foreach (var sbaRow in sbaRows)
            {
                var summary = new TestResultSummary();
                summary.AssemblyPathName = sbaRow.AssemblyPathName;
                summary.FullClassName = "";

                foreach (var ocn in outcomeNames)
                {
                    var oc = new ResultOutcomeSummary() { Outcome = ocn };
                    oc.NumTests = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Count();
                    oc.TotalDurationInSeconds = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Sum(r => r.DurationInSeconds);
                    summary.Outcomes.Add(oc);
                }

                summaries.Add(summary);
            }

            return summaries;
        }

        /// <summary>
        /// Summarises the test results by grouping by assembly and then by class.
        /// </summary>
        /// <param name="testResults">The test results.</param>
        /// <returns>Summary of results grouped by assembly and then by class.</returns>
        public static IEnumerable<TestResultSummary> SummariseByClass(IEnumerable<TestResult> testResults)
        {
            var summaries = new List<TestResultSummary>();
            var outcomeNames = ResultOutcomeSummary.GetOutcomeNames(testResults);

            var sbaRows = testResults.GroupBy(r => new { r.AssemblyPathName, r.ClassName })
                          .Select(gr => new
                          {
                              AssemblyPathName = gr.Key.AssemblyPathName,
                              FullClassName = gr.Key.ClassName,
                              TestResults = gr
                          })
                          .OrderBy(a => a.AssemblyPathName);

            foreach (var sbaRow in sbaRows)
            {
                var summary = new TestResultSummary();
                summary.AssemblyPathName = sbaRow.AssemblyPathName;
                summary.FullClassName = sbaRow.FullClassName;

                foreach (var ocn in outcomeNames)
                {
                    var oc = new ResultOutcomeSummary() { Outcome = ocn };
                    oc.NumTests = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Count();
                    oc.TotalDurationInSeconds = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Sum(r => r.DurationInSeconds);
                    summary.Outcomes.Add(oc);
                }

                summaries.Add(summary);
            }

            return summaries;
        }
    }
}
