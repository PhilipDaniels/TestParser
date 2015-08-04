using System.Globalization;
using System.IO;
using System.Text;

namespace TestParser.Core
{
    public class CSVTestResultWriter : ITestResultWriter
    {
        public void WriteResults(Stream s, TestResults testResults)
        {
            var utf8WithoutBom = new UTF8Encoding(false);

            using (var sw = new StreamWriter(s, utf8WithoutBom))
            {
                sw.Write("ResultsPathName,");
                sw.Write("ResultsFileName,");
                sw.Write("AssemblyPathName,");
                sw.Write("AssemblyFileName,");
                sw.Write("FullClassName,");
                sw.Write("ClassName,");
                sw.Write("TestName,");
                sw.Write("ComputerName,");
                sw.Write("StartTime,");
                sw.Write("EndTime,");
                sw.Write("Duration,");
                sw.Write("DurationInSeconds,");
                sw.Write("Outcome,");
                sw.Write("ErrorMessage,");
                sw.Write("StackTrace");
                sw.Write("TestResultFileType");
                sw.WriteLine();

                foreach (var r in testResults.Lines.SortedByFailedOtherPassed)
                {
                    sw.Write(Quoter.CSVQuote(r.ResultsPathName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.ResultsFileName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.AssemblyPathName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.AssemblyFileName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.FullClassName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.ClassName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.TestName)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.ComputerName)); sw.Write(",");

                    if (r.StartTime != null)
                    {
                        sw.Write(r.StartTime.Value.ToString("s")); sw.Write(",");
                    }
                    else
                    {
                        sw.Write(",");
                    }

                    if (r.EndTime != null)
                    {
                        sw.Write(r.EndTime.Value.ToString("s"));
                        sw.Write(",");
                    }
                    else
                    {
                        sw.Write(",");
                    }

                    sw.Write(r.DurationInSeconds.ToString("R", CultureInfo.InvariantCulture)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.Outcome)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.ErrorMessage)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.StackTrace)); sw.Write(",");
                    sw.Write(Quoter.CSVQuote(r.TestResultFileType.ToString()));
                    sw.WriteLine();
                }
            }
        }
    }
}
