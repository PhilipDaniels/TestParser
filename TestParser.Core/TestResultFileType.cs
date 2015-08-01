namespace TestParser.Core
{
    /// <summary>
    /// Possible file types.
    /// </summary>
    public enum TestResultFileType
    {
        /// <summary>
        /// The TRX file type, as produced by MS Test.
        /// </summary>
        Trx,

        /// <summary>
        /// The NUnit 2 file type, as produced by, er, NUnit 2.
        /// </summary>
        NUnit2,

            /// <summary>
        /// The NUnit 3 file type, as produced by, er, NUnit 3.
        /// </summary>
        NUnit3
    }
}
