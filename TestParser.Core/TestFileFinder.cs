using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TestParser.Core
{
    public class TestFileFinder
    {
        public IEnumerable<string> FindFiles(IEnumerable<string> globPatterns)
        {
            string cwd = Environment.CurrentDirectory;
            var actualFiles = new List<string>();

            foreach (string pattern in globPatterns)
            {
                if (Path.IsPathRooted(pattern))
                {
                    actualFiles.AddRange(Glob.Glob.ExpandNames(pattern));
                }
                else
                {
                    string p = Path.Combine(cwd, pattern);
                    actualFiles.AddRange(Glob.Glob.ExpandNames(p));
                }
            }

            actualFiles = actualFiles.Distinct().ToList();

            foreach (var f in actualFiles.OrderBy(f =>f))
            {
                if (File.Exists(f))
                    yield return f;
            }
        }
    }
}
