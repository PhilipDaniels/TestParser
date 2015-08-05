using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// A simple class to summarise a set of <see cref="TestResult"/>.
    /// </summary>
    [DebuggerDisplay("{AssemblyFileName}, Passed={TotalPassed}/{TotalTests} in {TotalDurationInSeconds} secs.")]
    public class TestResultSummaryLine : TestResultBase
    {
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
        /// Initializes a new instance of the <see cref="TestResultSummaryLine"/> class.
        /// </summary>
        public TestResultSummaryLine()
        {
            Outcomes = new List<ResultOutcomeSummary>();
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
                return TotalByOutcome(KnownOutcomes.Passed);
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

            return Outcomes.Where(c => c.Outcome == outcome).Single().NumTests;
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
    }
}
