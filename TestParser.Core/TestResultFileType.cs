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
        /// The NUnit file type, as produced by, er, NUnit.
        /// </summary>
        NUnit
    }
}
