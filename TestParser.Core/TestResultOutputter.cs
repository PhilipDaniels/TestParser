using System;
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
        /// <param name="cla">The command line arguments.</param>
        /// <exception cref="Exception">Output format must be specified.</exception>
        public void OutputResults(TestResults results, CommandLineArguments cla)
        {
            using (Stream s = cla.OutputFilename == null ? Console.OpenStandardOutput() : new FileStream(cla.OutputFilename, FileMode.Create))
            {
                ITestResultWriter writer = null;
                switch (cla.OutputFormat)
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
                        writer = new XLSXTestResultWriter(cla.YellowBand, cla.GreenBand);
                        break;
                    default:
                        throw new Exception("Unsupported output format: " + cla.OutputFormat);
                }

                writer.WriteResults(s, results);
            }
        }
    }
}
