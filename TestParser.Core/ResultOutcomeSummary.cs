using System.Collections.Generic;
using System.Diagnostics;

namespace TestParser.Core
{
    /// <summary>
    /// Represents the results of summarising test outcomes.
    /// </summary>
    [DebuggerDisplay("{Outcome}, Num={NumTests}, Duration={DurationInSeconds} secs.")]
    public class ResultOutcomeSummary
    {
        /// <summary>
        /// The name of the passed outcome.
        /// </summary>
        public const string PassedOutcome = "Passed";

        /// <summary>
        /// The name of the failed outcome.
        /// </summary>
        public const string FailedOutcome = "Failed";

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

        /// <summary>
        /// Gets all the outcome names contained within a set of test results.
        /// The names are guaranteed to include "Passed" and "Failed" as the first
        /// two items and any other outcomes will be sorted by name.
        /// </summary>
        /// <param name="testResults">The test results.</param>
        /// <returns>Sorted list of outcome names.</returns>
        public static IEnumerable<string> GetOutcomeNames(IEnumerable<TestResult> testResults)
        {
            var finalOutcomes = new List<string>();
            finalOutcomes.Add(PassedOutcome);
            finalOutcomes.Add(FailedOutcome);

            var remainingOutcomes = new List<string>();
            foreach (var r in testResults)
            {
                string oc = r.Outcome;
                if (oc != PassedOutcome && oc != FailedOutcome && !remainingOutcomes.Contains(oc))
                    remainingOutcomes.Add(oc);
            }

            finalOutcomes.AddRange(remainingOutcomes);

            return finalOutcomes;
        }
    }
}
