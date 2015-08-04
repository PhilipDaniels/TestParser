using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Store and summarise the entire set of test results.
    /// The summaries are mainly for convenience, equivalent or similar
    /// summaries can be created by Linq queries over the <see cref="Lines"/>.
    /// </summary>
    public class TestResults
    {
        List<string> outcomes; 
        
        public TestResultCollection Lines { get; private set; }
        public TestResultSummary SummaryByAssembly { get; private set; }
        public TestResultSummary SummaryByClass { get; private set; }
        public IEnumerable<string> Outcomes { get { return outcomes; } }

        public TestResults()
        {
            Lines = new TestResultCollection();
            outcomes = new List<string>();
        }

        public void Add(TestResult result)
        {
            result.ThrowIfNull("result");

            Lines.Add(result);
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
            outcomes = new List<string>();
            outcomes.Add(KnownOutcomes.Passed);
            outcomes.Add(KnownOutcomes.Failed);

            var remainingOutcomes = new List<string>();
            foreach (var result in Lines)
            {
                string oc = result.Outcome;
                if (oc != KnownOutcomes.Passed && oc != KnownOutcomes.Failed && !remainingOutcomes.Contains(oc))
                    remainingOutcomes.Add(oc);
            }

            remainingOutcomes.Sort();
            outcomes.AddRange(remainingOutcomes);
        }

        /// <summary>
        /// Summarises the test results by grouping by assembly.
        /// The <code>ClassName</code> property will be blank for these summary lines.
        /// </summary>
        void SummariseByAssembly()
        {
            SummaryByAssembly = new TestResultSummary();

            var sbaRows = Lines.GroupBy(r => r.AssemblyPathName)
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

                foreach (var ocn in Outcomes)
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

            var sbaRows = Lines.GroupBy(r => new { r.AssemblyPathName, r.ClassName })
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

                foreach (var ocn in Outcomes)
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
