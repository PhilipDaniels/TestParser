using System;
using TestParser.Core;

namespace TestParser
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var cla = new CommandLineArguments(args);
                if (cla.ErrorMessage != null)
                {
                    Console.Error.WriteLine(cla.GetUsageMessage());
                    Environment.Exit(1);
                }

                var resultsFactory = new TestResultFactory();
                var results = resultsFactory.CreateResultsFromTestFiles(cla.TestFiles);

                var outputCreator = new TestResultOutputter();
                outputCreator.OutputResults(results, cla.OutputFilename, cla.OutputFormat);
            }
            catch (Exception ex)
            {
                // Handy to put a breakpoint here :-)
                string msg = ex.Message;
                Console.Error.WriteLine(msg);
                Console.Error.WriteLine();
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}
