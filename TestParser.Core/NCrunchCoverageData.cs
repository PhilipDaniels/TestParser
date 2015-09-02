using System;
using System.Diagnostics;
using System.IO;

namespace TestParser.Core
{
    [DebuggerDisplay("{SourceFileName}, Coverage={Coverage}")]
    public class NCrunchCoverageData
    {
        /// <summary>
        /// Gets or sets the name of the project path.
        /// </summary>
        /// <value>
        /// The name of the project path.
        /// </value>
        public string ProjectPathName { get; set; }

        /// <summary>
        /// Gets the filename only of the <seealso cref="ProjectPathName"/>.
        /// </summary>
        /// <value>
        /// The name of the project file.
        /// </value>
        public string ProjectFileName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(ProjectPathName))
                    return "";
                else
                    return Path.GetFileName(ProjectPathName);
            }
        }

        /// <summary>
        /// Gets or sets the name of the source file path.
        /// </summary>
        /// <value>
        /// The name of the source file path.
        /// </value>
        public string SourceFilePathName { get; set; }

        /// <summary>
        /// Gets the filename only of the <seealso cref="SourceFilePathName"/>.
        /// </summary>
        /// <value>
        /// The name of the project file.
        /// </value>
        public string SourceFileName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(SourceFilePathName))
                    return "";
                else
                    return Path.GetFileName(SourceFilePathName);
            }
        }

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
    }
}
