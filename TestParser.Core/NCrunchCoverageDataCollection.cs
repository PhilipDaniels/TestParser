using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Holds a set of <see cref="NCrunchCoverageData"/> objects and provides
    /// some convenient sort methods.
    /// </summary>
    public class NCrunchCoverageDataCollection : IEnumerable<NCrunchCoverageData>
    {
        readonly List<NCrunchCoverageData> results = new List<NCrunchCoverageData>();

        /// <summary>
        /// Sort the results so that they are ordered by name.
        /// </summary>
        /// <returns>Ordered coverage data.</returns>
        public IEnumerable<NCrunchCoverageData> SortedByName
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
        public IEnumerable<NCrunchCoverageData> SortedByCoverage
        {
            get
            {
                return from r in results
                       orderby r.ProjectPathName, r.Coverage
                       select r;
            }
        }

        public IEnumerator<NCrunchCoverageData> GetEnumerator()
        {
            return results.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return results.GetEnumerator();
        }

        public void Add(NCrunchCoverageData coverageData)
        {
            coverageData.ThrowIfNull("coverageData");

            results.Add(coverageData);
        }

        public IEnumerable<NCrunchCoverageData> SummariseByProject
        {
            get
            {
                return from cd in this
                       group cd by cd.ProjectPathName into newGroup
                       orderby newGroup.Key
                       select new NCrunchCoverageData()
                       {
                           ProjectPathName = newGroup.Key,
                           CompiledLines = newGroup.Sum(g => g.CompiledLines),
                           CoveredLines = newGroup.Sum(g => g.CoveredLines)
                       };
            }
        }
    }
}
