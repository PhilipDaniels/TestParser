using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Holds a set of <see cref="CoverageData"/> objects and provides
    /// some convenient sort methods.
    /// </summary>
    public class CoverageDataCollection : IEnumerable<CoverageData>
    {
        readonly List<CoverageData> results = new List<CoverageData>();

        /// <summary>
        /// Sort the results so that they are ordered by name.
        /// </summary>
        /// <returns>Ordered coverage data.</returns>
        public IEnumerable<CoverageData> SortedByName
        {
            get
            {
                return from r in results
                       orderby r.ProjectPathName, r.SourceFilePathName
                       select r;
            }
        }

        /// <summary>
        /// Sort the results so that they are ordered by project name then coverage (worse is earlier).
        /// </summary>
        /// <returns>Ordered coverage data.</returns>
        public IEnumerable<CoverageData> SortedByCoverage
        {
            get
            {
                return from r in results
                       orderby r.ProjectPathName, r.Coverage
                       select r;
            }
        }

        public IEnumerator<CoverageData> GetEnumerator()
        {
            return results.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return results.GetEnumerator();
        }

        public void Add(CoverageData coverageData)
        {
            coverageData.ThrowIfNull("coverageData");

            results.Add(coverageData);
        }
    }
}
