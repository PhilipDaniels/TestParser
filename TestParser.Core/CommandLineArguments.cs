using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TestParser.Core
{
    /// <summary>
    /// A class to parse, or help with parsing, the command line arguments.
    /// </summary>
    public class CommandLineArguments
    {
        /// <summary>
        /// Gets the output format, as determined by the /of flag or the file extension.
        /// </summary>
        /// <value>
        /// The output format.
        /// </value>
        public OutputFormat OutputFormat { get; private set;  }

        /// <summary>
        /// Gets the output filename, if any.
        /// </summary>
        /// <value>
        /// The output filename.
        /// </value>
        public string OutputFilename { get; private set; }

        /// <summary>
        /// Gets the test files. These are the files to be parsed.
        /// This list is set by globbing the file parameters you specify on the
        /// command line.
        /// </summary>
        /// <value>
        /// The test files.
        /// </value>
        public List<string> TestFiles { get; private set; }

        /// <summary>
        /// Any remaining arguments. If there are switches which are not recognised
        /// by this class they are added to this list. This allows you to use this
        /// class to do the bulk of the work even if you are writing a custom driver
        /// program which accepts extra arguments.
        /// </summary>
        /// <value>
        /// The remaining arguments.
        /// </value>
        public List<string> RemainingArguments { get; private set; }

        /// <summary>
        /// Gets the error message. If this is non-null (even if it is String.Empty)
        /// you can assume that the usage message should be displayed and you should
        /// then exit.
        /// </summary>
        /// <value>
        /// The error message.
        /// </value>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandLineArguments"/> class.
        /// Parses and validates the known command line arguments. Any extra switches
        /// are ignored.
        /// </summary>
        /// <param name="args">The arguments.</param>
        public CommandLineArguments(string[] args)
        {
            OutputFormat = OutputFormat.Unspecified;
            TestFiles = new List<string>();
            RemainingArguments = new List<string>();

            if (args == null || args.Count() == 0)
            {
                ErrorMessage = "Error: No arguments specified.";
                return;
            }

            // Parse arguments.
            var fileGlobArguments = new List<string>();
            bool lookingAtFileArguments = false;
            foreach (string arg in args)
            {
                if (lookingAtFileArguments)
                {
                    fileGlobArguments.Add(arg);
                }
                else
                {
                    if (IsOptionArgument(arg))
                    {
                        ParseArgument(arg);
                    }
                    else
                    {
                        lookingAtFileArguments = true;
                        fileGlobArguments.Add(arg);
                    }
                }
            }

            // Validate them.
            if (OutputFilename != null && OutputFormat != OutputFormat.Unspecified)
            {
                ErrorMessage = "Error: If /of is specified /fmt cannot be used.";
                return;
            }

            if (OutputFilename == null && OutputFormat == OutputFormat.Unspecified)
            {
                OutputFormat = OutputFormat.CSV;
            }
            else
            {
                if (OutputFilename != null)
                {
                    string extension = Path.GetExtension(OutputFilename).Substring(1);
                    OutputFormat = GetFormat(extension);
                    if (OutputFormat == OutputFormat.Unspecified)
                        return;
                }
            }

            if (fileGlobArguments.Count == 0)
            {
                ErrorMessage = "Error: No test files specified.";
                return;
            }

            // Generate globbed list of test files to act as input.
            var finder = new TestFileFinder();
            TestFiles = finder.FindFiles(fileGlobArguments).ToList();

            if (TestFiles.Count == 0)
            {
                ErrorMessage = "Error: No test files found.";
                return;
            }
        }

        public string GetUsageMessage()
        {
            var sb = new StringBuilder();
            if (ErrorMessage != null)
                sb.AppendLine(ErrorMessage + Environment.NewLine + Environment.NewLine);

            sb.AppendLine("Usage");
            sb.AppendLine("=====");
            sb.AppendLine("TestParser.exe [/of:<filename> | /fmt:<format>] <inputfiles>");
            sb.AppendLine("");
            sb.AppendLine("If /fmt is specified, output is written to stdout. Valid formats");
            sb.AppendLine("are json, csv, kvp and xlsx.");
            sb.AppendLine("");
            sb.AppendLine("If /of is specified, output is written to a file and the format");
            sb.AppendLine("of the file is inferred from the extension. Valid extensions are");
            sb.AppendLine("the same as the /fmt argument.");

            return sb.ToString();
        }

        void ParseArgument(string arg)
        {
            if (arg == "/?" || arg == "/help")
            {
                // Will cause usage message to be displayed.
                ErrorMessage = "";
                return;
            }

            if (arg.StartsWith("/fmt:"))
            {
                OutputFormat = GetFormat(arg.Substring(5));
                return;
            }

            if (arg.StartsWith("/of:"))
            {
                OutputFilename = arg.Substring(4);
                if (String.IsNullOrWhiteSpace(OutputFilename))
                {
                    ErrorMessage = "Error: No filename specified for /of.";
                }
                return;
            }

            RemainingArguments.Add(arg);
        }

        bool IsOptionArgument(string arg)
        {
            return Regex.IsMatch(arg, @"/\w*:\w*");
        }

        OutputFormat GetFormat(string fmt)
        {
            if (fmt.Equals("csv", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.CSV;
            else if (fmt.Equals("json", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.Json;
            else if (fmt.Equals("kvp", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.KVP;
            else if (fmt.Equals("xlsx", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.Xlsx;
            
            ErrorMessage = "Error: Unknown format: " + fmt;

            return OutputFormat.Unspecified;
        }
    }
}
