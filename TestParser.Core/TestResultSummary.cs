using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    public class TestResultSummary
    {
        private readonly List<TestResultSummaryLine> summaryLines;

        public IEnumerable<TestResultSummaryLine> SummaryLines { get { return summaryLines; } }

        public TestResultSummary()
        {
            summaryLines = new List<TestResultSummaryLine>();
        }

        public void Add(TestResultSummaryLine summaryLine)
        {
            summaryLine.ThrowIfNull("summaryLine");

            summaryLines.Add(summaryLine);
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
                return SummaryLines.Sum(c => c.TotalTests);
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
                return SummaryLines.Sum(c => c.TotalPassed);
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

            return SummaryLines.Sum(c => c.TotalByOutcome(outcome));
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
                return SummaryLines.Sum(c => c.TotalDurationInSeconds);
            }
        }

        /// <summary>
        /// Gets the total duration in seconds as a human readable string of the form "hh:mm:ss.ff".
        /// </summary>
        /// <value>
        /// The total duration in seconds as a human readable string of the form "hh:mm:ss.ff".
        /// </value>
        public string TotalDurationInSecondsHuman
        {
            get
            {
                return HumanTime.ToHumanString(TotalDurationInSeconds);
            }
        }
    }
}
