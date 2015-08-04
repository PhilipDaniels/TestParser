using System.Diagnostics;

namespace TestParser.Core
{
    /// <summary>
    /// Represents the results of summarising tests by outcome.
    /// </summary>
    [DebuggerDisplay("{Outcome}, Num={NumTests}, Duration={DurationInSeconds} secs.")]
    public class ResultOutcomeSummary
    {
        /// <summary>
        /// Gets or sets the outcome, e.g. "Passed" or "Inconclusive".
        /// </summary>
        /// <value>
        /// The outcome.
        /// </value>
        public string Outcome { get; set; }

        /// <summary>
        /// Gets or sets the number tests.
        /// </summary>
        /// <value>
        /// The number tests.
        /// </value>
        public int NumTests { get; set; }

        /// <summary>
        /// Gets or sets the total duration in seconds.
        /// </summary>
        /// <value>
        /// The duration in seconds.
        /// </value>
        public double TotalDurationInSeconds { get; set; }
    }
}
