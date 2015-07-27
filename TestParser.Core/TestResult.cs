using System;
using System.Diagnostics;
using System.IO;

namespace TestParser.Core
{
    [DebuggerDisplay("{TestName}")]
    public class TestResult
    {
        public string ResultsPathName { get; set; }
        public string AssemblyPathName { get; set; }
        public string FullClassName { get; set; }
        public string TestName { get; set; }

        public string ComputerName { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }

        public string Outcome { get; set; }
        public string ErrorMessage { get; set; }
        public string StackTrace { get; set; }

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

        public string AssemblyFileName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(AssemblyPathName))
                    return "";
                else
                    return Path.GetFileName(AssemblyPathName);
            }
        }

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

        public string ClassName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(FullClassName))
                    return "";

                int idx = FullClassName.IndexOf(", ");
                if (idx > 0)
                    return FullClassName.Substring(0, idx);
                else
                    return FullClassName;
            }
        }

        public TimeSpan Duration
        {
            get
            {
                return EndTime - StartTime;
            }
        }

        public double DurationInSeconds
        {
            get
            {
                return Duration.TotalSeconds;
            }
        }
    }
}
