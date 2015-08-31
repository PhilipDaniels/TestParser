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

        /*
        /// <summary>
        /// Gets or sets the number of compiled lines.
        /// </summary>
        /// <value>
        /// The number of compiled lines.
        /// </value>
        public int CompiledLines { get; set; }

        /// <summary>
        /// Gets or sets the number of covered lines.
        /// </summary>
        /// <value>
        /// The number of covered lines.
        /// </value>
        public int CoveredLines { get; set; }

        /// <summary>
        /// Gets the number of uncovered lines.
        /// </summary>
        /// <value>
        /// The number of uncovered lines.
        /// </value>
        public int UncoveredLines
        {
            get
            {
                return CompiledLines - CoveredLines;
            }
        }

        /// <summary>
        /// Gets the coverage percentage, expressed as a number 0..1.
        /// </summary>
        /// <value>
        /// The coverage percentage.
        /// </value>
        public double Coverage
        {
            get
            {
                if (CompiledLines == 0)
                    return 0;
                else
                    return (double)CoveredLines / (double)CompiledLines;
            }
        }
        */
    }
}
