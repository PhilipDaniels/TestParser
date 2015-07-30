using System;
using System.Diagnostics;
using System.IO;

namespace TestParser.Core
{
    /// <summary>
    /// Represents the result of running a single test.
    /// Created based on the trx output format of MS Test.
    /// </summary>
    [DebuggerDisplay("{TestName}")]
    public class TestResult : TestResultBase
    {
        /// <summary>
        /// Gets or sets the name of the results path.
        /// </summary>
        /// <value>
        /// The name of the results path.
        /// </value>
        public string ResultsPathName { get; set; }

        /// <summary>
        /// Gets or sets the name of the test.
        /// </summary>
        /// <value>
        /// The name of the test.
        /// </value>
        public string TestName { get; set; }

        /// <summary>
        /// Gets or sets the name of the computer on which the test was run.
        /// </summary>
        /// <value>
        /// The name of the computer.
        /// </value>
        public string ComputerName { get; set; }

        /// <summary>
        /// Gets or sets the start time of the test. UTC.
        /// </summary>
        /// <value>
        /// The start time.
        /// </value>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end time of the test. UTC.
        /// </summary>
        /// <value>
        /// The end time.
        /// </value>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets the outcome. The outcomes "Passed" and "Failed"
        /// are "well known names" which are used in various summary
        /// routines.
        /// </summary>
        /// <value>
        /// The outcome.
        /// </value>
        public string Outcome { get; set; }

        /// <summary>
        /// Gets or sets the error message. Only set if the test failed.
        /// </summary>
        /// <value>
        /// The error message.
        /// </value>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the stack trace. Only set if the test failed.
        /// </summary>
        /// <value>
        /// The stack trace.
        /// </value>
        public string StackTrace { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TestResult"/> class.
        /// Sets string properties to empty string for ease of exporting/summarising.
        /// </summary>
        public TestResult()
        {
            ResultsPathName = "";
            AssemblyPathName = "";
            FullClassName = "";
            TestName = "";
            ComputerName = "";
            StartTime = DateTime.MinValue;
            EndTime = DateTime.MinValue;
            Outcome = "";
            ErrorMessage = "";
            StackTrace = "";
        }

        /// <summary>
        /// Gets the name of the results file without its path.
        /// </summary>
        /// <value>
        /// The name of the results file.
        /// </value>
        public string ResultsFileName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(ResultsPathName))
                    return "";
                else
                    return Path.GetFileName(ResultsPathName);
            }
        }

        /// <summary>
        /// Gets the duration of the test, which is the difference between
        /// the <see cref="EndTime"/> and the <see cref="StartTime"/>.
        /// </summary>
        /// <value>
        /// The duration.
        /// </value>
        public TimeSpan Duration
        {
            get
            {
                return EndTime - StartTime;
            }
        }

        /// <summary>
        /// Gets the duration in seconds.
        /// </summary>
        /// <value>
        /// The duration in seconds.
        /// </value>
        public double DurationInSeconds
        {
            get
            {
                return Duration.TotalSeconds;
            }
        }
    }
}
