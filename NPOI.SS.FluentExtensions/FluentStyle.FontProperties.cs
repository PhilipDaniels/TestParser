using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Provides a data structure which can be used to accumulate
    /// styling options via a Fluent API, until they need to be applied
    /// to a cell.
    /// </summary>
    public partial class FluentStyle
    {
        /// <summary>
        /// Gets or sets the font weight.
        /// </summary>
        /// <value>
        /// The font weight.
        /// </value>
        public FontBoldWeight? FontWeight { get; set; }

        /// <summary>
        /// Gets or sets the charset.
        /// </summary>
        /// <value>
        /// The charset.
        /// </value>
        public short? Charset { get; set; }

        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        /// <value>
        /// The color.
        /// </value>
        public short? Color { get; set; }

        /// <summary>
        /// Gets or sets the height of the font.
        /// </summary>
        /// <value>
        /// The height of the font.
        /// </value>
        public double? FontHeight { get; set; }

        /// <summary>
        /// Gets or sets the font height in points.
        /// </summary>
        /// <value>
        /// The font height in points.
        /// </value>
        public short? FontHeightInPoints { get; set; }

        /// <summary>
        /// Gets or sets the name of the font.
        /// </summary>
        /// <value>
        /// The name of the font.
        /// </value>
        public string FontName { get; set; }

        /// <summary>
        /// Gets or sets the italic setting.
        /// </summary>
        /// <value>
        /// The italic setting.
        /// </value>
        public bool? Italic { get; set; }

        /// <summary>
        /// Gets or sets the strikeout setting.
        /// </summary>
        /// <value>
        /// The strikeout setting.
        /// </value>
        public bool? Strikeout { get; set; }

        /// <summary>
        /// Gets or sets the super script setting.
        /// </summary>
        /// <value>
        /// The super script setting.
        /// </value>
        public FontSuperScript? SuperScript { get; set; }

        /// <summary>
        /// Gets or sets the underline setting.
        /// </summary>
        /// <value>
        /// The underline setting.
        /// </value>
        public FontUnderlineType? Underline { get; set; }
    }
}
