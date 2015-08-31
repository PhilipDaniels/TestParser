using System.Collections.Generic;
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
