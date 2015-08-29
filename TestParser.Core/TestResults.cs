using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Store and summarise the entire set of test results.
    /// The summaries are mainly for convenience, equivalent or similar
    /// summaries can be created by Linq queries over the <see cref="ResultLines"/>.
    /// </summary>
    public class TestResults
    {
        List<string> outcomeNames;

        public TestResultCollection ResultLines { get; private set; }
        public TestResultSummary SummaryByAssembly { get; private set; }
        public TestResultSummary SummaryByClass { get; private set; }
        public IEnumerable<string> OutcomeNames { get { return outcomeNames; } }

        /// <summary>
        /// Gets the 10 slowest tests.
        /// </summary>
        /// <value>
        /// The 10 slowest tests.
        /// </value>
        public IEnumerable<SlowestTest> SlowestTests
        {
            get
            {
                return ResultLines.OrderByDescending(r => r.DurationInSeconds)
                                  .Take(10)
                                  .Select(r => new SlowestTest()
                                  {
                                      AssemblyPathName = r.AssemblyPathName,
                                      DurationInSeconds = r.DurationInSeconds,
                                      FullClassName = r.FullClassName,
                                      TestName = r.TestName
                                  });
            }
        }

        public TestResults()
        {
            ResultLines = new TestResultCollection();
            outcomeNames = new List<string>();
        }

        public void Add(TestResult result)
        {
            result.ThrowIfNull("result");

            ResultLines.Add(result);
        }

        public void AddRange(IEnumerable<TestResult> results)
        {
            results.ThrowIfNull("results");

            foreach (var result in results)
                Add(result);
        }

        /// <summary>
        /// Summarises this instance. Call this when you have finished adding lines.
        /// </summary>
        public void Summarise()
        {
            SummariseOutcomeNames();
            SummariseByAssembly();
            SummariseByClass();
        }

        /// <summary>
        /// Summarise all the outcome names contained within a set of test results.
        /// The names are guaranteed to include "Passed" and "Failed" as the first
        /// two items and any other outcomes will be sorted by name.
        /// </summary>
        void SummariseOutcomeNames()
        {
            outcomeNames = new List<string>();
            outcomeNames.Add(KnownOutcomes.Passed);
            outcomeNames.Add(KnownOutcomes.Failed);

            var remainingOutcomes = new List<string>();
            foreach (var result in ResultLines)
            {
                string oc = result.Outcome;
                if (oc != KnownOutcomes.Passed && oc != KnownOutcomes.Failed && !remainingOutcomes.Contains(oc))
                    remainingOutcomes.Add(oc);
            }

            remainingOutcomes.Sort();
            outcomeNames.AddRange(remainingOutcomes);
        }

        /// <summary>
        /// Summarises the test results by grouping by assembly.
        /// The <code>ClassName</code> property will be blank for these summary lines.
        /// </summary>
        void SummariseByAssembly()
        {
            SummaryByAssembly = new TestResultSummary();

            var sbaRows = ResultLines.GroupBy(r => r.AssemblyPathName)
                          .Select(gr => new
                          {
                              AssemblyPathName = gr.Key,
                              TestResults = gr
                          });

            foreach (var row in sbaRows)
            {
                var summary = new TestResultSummaryLine();
                summary.AssemblyPathName = row.AssemblyPathName;
                summary.FullClassName = "";

                foreach (var ocn in OutcomeNames)
                {
                    var oc = new ResultOutcomeSummary() { Outcome = ocn };
                    oc.NumTests = (from r in row.TestResults where r.Outcome == ocn select r).Count();
                    oc.TotalDurationInSeconds = (from r in row.TestResults where r.Outcome == ocn select r).Sum(r => r.DurationInSeconds);
                    summary.Outcomes.Add(oc);
                }

                SummaryByAssembly.Add(summary);
            }
        }

        /// <summary>
        /// Summarises the test results by grouping by assembly and then by class.
        /// </summary>
        void SummariseByClass()
        {
            SummaryByClass = new TestResultSummary();

            var sbaRows = ResultLines.GroupBy(r => new { r.AssemblyPathName, r.ClassName })
                          .Select(gr => new
                          {
                              AssemblyPathName = gr.Key.AssemblyPathName,
                              FullClassName = gr.Key.ClassName,
                              TestResults = gr
                          });

            foreach (var sbaRow in sbaRows)
            {
                var summary = new TestResultSummaryLine();
                summary.AssemblyPathName = sbaRow.AssemblyPathName;
                summary.FullClassName = sbaRow.FullClassName;

                foreach (var ocn in OutcomeNames)
                {
                    var oc = new ResultOutcomeSummary() { Outcome = ocn };
                    oc.NumTests = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Count();
                    oc.TotalDurationInSeconds = (from r in sbaRow.TestResults where r.Outcome == ocn select r).Sum(r => r.DurationInSeconds);
                    summary.Outcomes.Add(oc);
                }

                SummaryByClass.Add(summary);
            }
        }
    }
}
