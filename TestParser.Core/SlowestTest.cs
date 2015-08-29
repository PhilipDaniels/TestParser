namespace TestParser.Core
{
    /// <summary>
    /// Represents information about the slowest tests.
    /// Exists so we can easily incorporate that into the summary.
    /// </summary>
    public class SlowestTest : TestResultBase
    {
        /// <summary>
        /// Gets or sets the name of the test.
        /// </summary>
        /// <value>
        /// The name of the test.
        /// </value>
        public string TestName { get; set; }

        /// <summary>
        /// Gets or sets the duration in seconds.
        /// </summary>
        /// <value>
        /// The duration in seconds.
        /// </value>
        public double DurationInSeconds { get; set; }

        /// <summary>
        /// Returns the duration in a human-readable "hh:mm:ss.ff" format.
        /// </summary>
        /// <value>
        /// The duration in a human-readable "hh:mm:ss.ff" format.
        /// </value>
        public string DurationHuman
        {
            get
            {
                return HumanTime.ToHumanString(DurationInSeconds);
            }
        }
    }
}
