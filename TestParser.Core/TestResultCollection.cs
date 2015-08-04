using System.Collections.Generic;
using System.Linq;
using BassUtils;

namespace TestParser.Core
{
    /// <summary>
    /// Holds a set of <see cref="TestResult"/> objects and provides
    /// some convenient sort methods.
    /// </summary>
    public class TestResultCollection : IEnumerable<TestResult>
    {
        readonly List<TestResult> results = new List<TestResult>();

        /// <summary>
        /// Sort the results so that they are ordered: "Failures, NonPassed, Passed".
        /// This puts the most attention-needing results first.
        /// </summary>
        /// <returns>Ordered results.</returns>
        public IEnumerable<TestResult> SortedByFailedOtherPassed
        {
            get
            {
                return from r in results
                       let sortGroup = r.Outcome == KnownOutcomes.Passed ? 2 :
                                       r.Outcome == KnownOutcomes.Failed ? 0 : 1
                       orderby sortGroup, r.ResultsPathName, r.AssemblyPathName, r.ClassName, r.TestName
                       select r;
            }
        }

        public IEnumerable<TestResult> SortedByFileAndClass
        {
            get
            {
                return from r in results
                       orderby r.ResultsPathName, r.AssemblyPathName, r.ClassName, r.TestName
                       select r;
            }
        }

        public IEnumerator<TestResult> GetEnumerator()
        {
            return results.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return results.GetEnumerator();
        }

        public void Add(TestResult result)
        {
            result.ThrowIfNull("result");

            results.Add(result);
        }
    }
}
