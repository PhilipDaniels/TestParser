using System.Collections.Generic;
using System.Diagnostics;

namespace TestParser.Core
{
    [DebuggerDisplay("{Outcome}, Num={NumTests}, Duration={DurationInSeconds} secs.")]
    public class ResultOutcome
    {
        public const string PassedOutcome = "Passed";
        public const string FailedOutcome = "Failed";

        public string Outcome { get; set; }
        public int NumTests { get; set; }
        public double DurationInSeconds { get; set; }

        public static IEnumerable<string> GetOutcomeNames(IEnumerable<TestResult> testResults)
        {
            var finalOutcomes = new List<string>();
            finalOutcomes.Add(PassedOutcome);
            finalOutcomes.Add(FailedOutcome);

            foreach (var r in testResults)
            {
                if (!finalOutcomes.Contains(r.Outcome))
                    finalOutcomes.Add(r.Outcome);
            }

            return finalOutcomes;
        }
    }
}
