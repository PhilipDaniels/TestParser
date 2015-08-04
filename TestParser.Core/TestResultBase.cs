using System;
using System.IO;

namespace TestParser.Core
{
    /// <summary>
    /// Contains some properties that are common to a <see cref="TestResult"/>
    /// and a <see cref="TestResultSummaryLine"/>.
    /// </summary>
    public abstract class TestResultBase
    {
        /// <summary>
        /// The full path of the assembly that contained the tests.
        /// </summary>
        public string AssemblyPathName { get; set; }

        /// <summary>
        /// The filename of the assembly that contained the tests.
        /// </summary>
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

        /// <summary>
        /// The fully-qualified name of the class that contained the tests.
        /// </summary>
        public string FullClassName { get; set; }

        /// <summary>
        /// The partial name of the class that contained the tests, as entered
        /// into the source code editor.
        /// </summary>
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
    }
}
