﻿using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace TestParser.Core
{
    public class KVPTestResultWriter : ITestResultWriter
    {
        public void WriteResults(Stream s, IEnumerable<TestResult> testResults)
        {
            var utf8WithoutBom = new UTF8Encoding(false);

            using (var sw = new StreamWriter(s, utf8WithoutBom))
            {
                foreach (var r in testResults)
                {
                    sw.Write(" ResultsPathName=" + Quoter.KVPQuote(r.ResultsPathName));
                    sw.Write(" ResultsFileName=" + Quoter.KVPQuote(r.ResultsFileName));
                    sw.Write(" AssemblyPathName=" + Quoter.KVPQuote(r.AssemblyPathName));
                    sw.Write(" AssemblyFileName=" + Quoter.KVPQuote(r.AssemblyFileName));
                    sw.Write(" FullClassName=" + Quoter.KVPQuote(r.FullClassName));
                    sw.Write(" ClassName=" + Quoter.KVPQuote(r.ClassName));
                    sw.Write(" TestName=" + Quoter.KVPQuote(r.TestName));
                    sw.Write(" ComputerName=" + Quoter.KVPQuote(r.ComputerName));
                    sw.Write(" StartTime=" + Quoter.KVPQuote(r.StartTime.ToString("s")));
                    sw.Write(" EndTime=" + Quoter.KVPQuote(r.EndTime.ToString("s")));
                    sw.Write(" Duration=" + Quoter.KVPQuote(r.Duration.ToString("c")));
                    sw.Write(" DurationInSeconds=" + Quoter.KVPQuote(r.DurationInSeconds.ToString("R", CultureInfo.InvariantCulture)));
                    sw.Write(" Outcome=" + Quoter.KVPQuote(r.Outcome));
                    sw.Write(" ErrorMessage=" + Quoter.KVPQuote(r.ErrorMessage));
                    sw.Write(" StackTrace=" + Quoter.KVPQuote(r.StackTrace));
                    sw.WriteLine();
                }
            }
        }
    }
}
