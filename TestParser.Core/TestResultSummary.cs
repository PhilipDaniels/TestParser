using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    public class TestResultSummary
    {
        private readonly List<TestResultSummaryLine> lines;

        public IEnumerable<TestResultSummaryLine> Lines { get { return lines; } }

        public TestResultSummary()
        {
            lines = new List<TestResultSummaryLine>();
        }

        public void Add(TestResultSummaryLine summaryLine)
        {
            summaryLine.ThrowIfNull("summaryLine");

            lines.Add(summaryLine);
        }

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
                return Lines.Sum(c => c.TotalTests);
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
                return Lines.Sum(c => c.TotalPassed);
            }
        }

        /// <summary>
        /// Gets the total number of tests by outcome.
        /// </summary>
        /// <param name="outcome">The outcome.</param>
        /// <returns>Number of tests.</returns>
        public int TotalByOutcome(string outcome)
        {
            outcome.ThrowIfNull("outcome");

            return Lines.Sum(c => c.TotalByOutcome(outcome));
        }

        /// <summary>
        /// Gets the percent passed (0..1).
        /// </summary>
        /// <value>
        /// The percent passed.
        /// </value>
        public double PercentPassed
        {
            get
            {
                if (TotalTests == 0)
                    return 0;
                else
                    return (double)TotalPassed / (double)TotalTests;
            }
        }

        /// <summary>
        /// Gets the total duration in seconds of this summary set.
        /// </summary>
        /// <value>
        /// The total duration in seconds of this summary set.
        /// </value>
        public double TotalDurationInSeconds
        {
            get
            {
                return Lines.Sum(c => c.TotalDurationInSeconds);
            }
        }
    }
}
