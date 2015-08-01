namespace TestParser.Core
{
    /// <summary>
    /// An enumeration representing the different output formats the program supports.
    /// </summary>
    public enum OutputFormat
    {
        /// <summary>
        /// The unspecified format.
        /// </summary>
        Unspecified,

        /// <summary>
        /// The CSV format.
        /// </summary>
        CSV,

        /// <summary>
        /// The Json format.
        /// </summary>
        Json,

        /// <summary>
        /// The KVP format, produces output like "TestName=FooTest, ..." on each line.
        /// </summary>
        KVP,

        /// <summary>
        /// Excel spreadsheet.
        /// </summary>
        Xlsx
    }
}
