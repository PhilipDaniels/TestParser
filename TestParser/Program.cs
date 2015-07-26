using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TestParser.Core;

namespace TestParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileArguments = new List<string>();
            OutputFormat outputFormat = OutputFormat.Unspecified;
            string outputFilename = null;

            bool lookingAtFileArguments = false;
            foreach (string arg in args)
            {
                if (lookingAtFileArguments)
                {
                    fileArguments.Add(arg);
                }
                else
                {
                    if (arg.StartsWith("/?"))
                    {
                        ShowUsage();
                        Environment.Exit(1);
                    }
                    else if (arg.StartsWith("/fmt:"))
                    {
                        outputFormat = GetFormat(arg.Substring(5));
                    }
                    else if (arg.StartsWith("/of:"))
                    {
                        outputFilename = arg.Substring(4);
                    }
                    else
                    {
                        lookingAtFileArguments = true;
                        fileArguments.Add(arg);
                    }
                }
            }

            if (outputFilename != null && outputFormat != OutputFormat.Unspecified)
            {
                Console.Error.WriteLine("Error: If /of is specified /fmt cannot be used.");
                ShowUsage();
                Environment.Exit(1);
            }

            if (outputFilename == null && outputFormat == OutputFormat.Unspecified)
            {
                outputFormat = OutputFormat.CSV;
            }
            else
            {
                if (outputFilename != null)
                    outputFormat = GetFormatFromFilename(outputFilename);
            }

            if (fileArguments.Count == 0)
            {
                Console.Error.WriteLine("Error: No test files specified.");
                ShowUsage();
                Environment.Exit(1);
            }

            fileArguments = GlobFileArguments(fileArguments);

            if (fileArguments.Count == 0)
            {
                Console.Error.WriteLine("Error: No test files found.");
                ShowUsage();
                Environment.Exit(1);
            }


            var results = GetResults(fileArguments);
            OutputResults(results, outputFilename, outputFormat);
        }

        static List<string> GlobFileArguments(IEnumerable<string> fileArguments)
        {
            var finder = new TestFileFinder();
            return finder.FindFiles(fileArguments).ToList();
        }

        static OutputFormat GetFormat(string arg)
        {
            if (arg.Equals("csv", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.CSV;
            else if (arg.Equals("json", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.Json;
            else if (arg.Equals("kvp", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.KVP;

            Console.Error.WriteLine("Error: Invalid /fmt:" + arg);
            ShowUsage();
            Environment.Exit(1);

            // Get the compiler to shut up about return paths.
            return OutputFormat.Unspecified;
        }

        static OutputFormat GetFormatFromFilename(string outputFilename)
        {
            string ext = Path.GetExtension(outputFilename).Substring(1);

            if (ext.Equals("csv", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.CSV;
            else if (ext.Equals("json", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.Json;
            else if (ext.Equals("kvp", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.KVP;
            else if (ext.Equals("xlsx", StringComparison.InvariantCultureIgnoreCase))
                return OutputFormat.Xlsx;

            Console.Error.WriteLine("Error: Invalid file extension, cannot infer type: " + outputFilename);
            ShowUsage();
            Environment.Exit(1);

            // Get the compiler to shut up about return paths.
            return OutputFormat.Unspecified;
        }

        static void ShowUsage()
        {
            Console.Error.WriteLine("");
            Console.Error.WriteLine("Usage");
            Console.Error.WriteLine("=====");
            Console.Error.WriteLine("TestParser.exe [/of:<filename> | /fmt:<format>] <inputfiles>");
            Console.Error.WriteLine("");
            Console.Error.WriteLine("If /fmt is specified, output is written to stdout. Valid formats");
            Console.Error.WriteLine("are json, csv and kvp.");
            Console.Error.WriteLine("");
            Console.Error.WriteLine("If /of is specified, output is written to a file and the format");
            Console.Error.WriteLine("of the file is inferred from the extension. Valid extensions are");
            Console.Error.WriteLine("json, csv, kvp and xlsx.");
        }

        static IEnumerable<TestResult> GetResults(IEnumerable<string> fileArguments)
        {
            var results = new List<TestResult>();
            var parser = new TrxFileParser();
            foreach (string file in fileArguments)
            {
                results.AddRange(parser.Parse(file));
            }

            return results;
        }

        static void OutputResults(IEnumerable<TestResult> results, string outputFilename, OutputFormat outputFormat)
        {
            using (Stream s = outputFilename == null ? Console.OpenStandardOutput() : new FileStream(outputFilename, FileMode.Create))
            {
                ITestResultWriter writer = null;
                switch (outputFormat)
                {
                    case OutputFormat.Unspecified:
                        throw new Exception("Output format must be specified.");
                    case OutputFormat.CSV:
                        writer = new CSVTestResultWriter();
                        break;
                    case OutputFormat.Json:
                        writer = new JSONTestResultWriter();
                        break;
                    case OutputFormat.KVP:
                        writer = new KVPTestResultWriter();
                        break;
                    case OutputFormat.Xlsx:
                        writer = new XLSXTestResultWriter();
                        break;
                    default:
                        break;
                }

                writer.WriteResults(s, results);
            }



            //StreamWriter sw = null;
            //if (outputFilename == null)
            //{
            //    sw = new StreamWriter(Console.OpenStandardOutput());
            //    sw.AutoFlush = true;
            //    Console.SetOut(sw);
            //}
            //else
            //{
            //    sw = new StreamWriter(outputFilename, false, Encoding.UTF8);
            //}

            //sw.Write("hello world");
        }

    }
}
