using System;
using System.Collections.Generic;
using System.IO;

namespace TestParser.Core
{
    /// <summary>
    /// Simple wrapper class to generate the appropriate type of result stream.
    /// </summary>
    public class TestResultOutputter
    {
        /// <summary>
        /// Outputs the results to an appropriate file or stdout.
        /// </summary>
        /// <param name="results">The test results.</param>
        /// <param name="outputFilename">The output filename. If null, output is to stdout.</param>
        /// <param name="outputFormat">The output format.</param>
        /// <exception cref="Exception">Output format must be specified.</exception>
        public void OutputResults(IEnumerable<TestResult> results, string outputFilename, OutputFormat outputFormat)
        {
            using (Stream s = outputFilename == null ? Console.OpenStandardOutput() : new FileStream(outputFilename, FileMode.Create))
            {
                ITestResultWriter writer = null;
                switch (outputFormat)
                {
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
                        throw new Exception("Unsupported output format: " + outputFormat);
                }

                writer.WriteResults(s, results);
            }
        }
    }
}
