using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Bundle together an <c>ICell</c> and a <see cref="FluentStyle"/>
    /// so that they the cell can be configured with a fluent interface.
    /// </summary>
    public partial class FluentCell
    {
        /// <summary>
        /// Set the font weight.
        /// </summary>
        /// <param name="fontWeight">The font weight.</param>
        /// <returns>The cell.</returns>
        public FluentCell FontWeight(FontBoldWeight fontWeight)
        {
            Style.FontWeight = fontWeight;
            return this;
        }

        /// <summary>
        /// Set the charset.
        /// </summary>
        /// <param name="charset">The charset.</param>
        /// <returns>The cell.</returns>
        public FluentCell Charset(short charset)
        {
            Style.Charset = charset;
            return this;
        }

        /// <summary>
        /// Sets the color.
        /// </summary>
        /// <param name="color">The color.</param>
        /// <returns>The cell.</returns>
        public FluentCell Color(short color)
        {
            Style.Color = color;
            return this;
        }

        /// <summary>
        /// Sets the font height.
        /// </summary>
        /// <param name="fontHeight">Height of the font.</param>
        /// <returns>The cell.</returns>
        public FluentCell FontHeight(double fontHeight)
        {
            Style.FontHeight = fontHeight;
            return this;
        }

        /// <summary>
        /// Sets the font height in points.
        /// </summary>
        /// <param name="fontHeightInPoints">The font height in points.</param>
        /// <returns>The cell.</returns>
        public FluentCell FontHeightInPoints(short fontHeightInPoints)
        {
            Style.FontHeightInPoints = fontHeightInPoints;
            return this;
        }

        /// <summary>
        /// Sets the font name.
        /// </summary>
        /// <param name="fontName">Name of the font.</param>
        /// <returns>The cell.</returns>
        public FluentCell FontName(string fontName)
        {
            Style.FontName = fontName;
            return this;
        }

        /// <summary>
        /// Sets italics on or off.
        /// </summary>
        /// <param name="italic">if set to <c>true</c> [italic].</param>
        /// <returns>The cell.</returns>
        public FluentCell Italic(bool italic)
        {
            Style.Italic = italic;
            return this;
        }

        /// <summary>
        /// Sets strikeout on or off.
        /// </summary>
        /// <param name="strikeout">if set to <c>true</c> [strikeout].</param>
        /// <returns>The cell.</returns>
        public FluentCell Strikeout(bool strikeout)
        {
            Style.Strikeout = strikeout;
            return this;
        }

        /// <summary>
        /// Sets the superscript type.
        /// </summary>
        /// <param name="superScript">The super script.</param>
        /// <returns>The cell.</returns>
        public FluentCell SuperScript(FontSuperScript superScript)
        {
            Style.SuperScript = superScript;
            return this;
        }

        /// <summary>
        /// Sets the underline type.
        /// </summary>
        /// <param name="underline">The underline.</param>
        /// <returns>The cell.</returns>
        public FluentCell Underline(FontUnderlineType underline)
        {
            Style.Underline = underline;
            return this;
        }
    }
}
