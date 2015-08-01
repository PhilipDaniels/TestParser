using System;
using System.IO;

namespace TestParser.Core
{
    /// <summary>
    /// Try and figure out the type of data contained in a file.
    /// </summary>
    public static class TestResultFileTypeGuesser
    {
        /// <summary>
        /// Guesses the type of the file.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns></returns>
        /// <exception cref="System.IO.FileNotFoundException">The file ' + filename + ' does not exist.</exception>
        public static TestResultFileType GuessFileType(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException("The file '" + filename + "' does not exist.", filename);

            string extension = Path.GetExtension(filename).Substring(1);
            if (extension.Equals("trx", StringComparison.InvariantCultureIgnoreCase))
                return TestResultFileType.Trx;

            // NUnit by default creates files ending in ".xml" but that is not guaranteed.


            // Still no luck? Open the file and look for some stuff.
            string text = File.ReadAllText(filename);
            if (text.Contains("xmlns=\"http://microsoft.com/schemas/VisualStudio/TeamTest/2010\""))
                return TestResultFileType.Trx;


            throw new Exception("Could not guess type of data in " + filename);
        }
    }
}
