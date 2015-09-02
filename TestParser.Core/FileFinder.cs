using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TestParser.Core
{
    /// <summary>
    /// Finds the unique set of files matching a set of glob patterns.
    /// For relative paths, the globbing is done relative to the current working directory.
    /// The returned set may be empty.
    /// </summary>
    public class FileFinder
    {
        /// <summary>
        /// Find the unique set of files that match the set of <paramref name="globPatterns."/>
        /// </summary>
        /// <param name="globPatterns">The glob patterns. ** wildcards may be used, and the
        /// patterns may be relative to the current working directory or absolute.</param>
        /// <returns>Set of matching files.</returns>
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
