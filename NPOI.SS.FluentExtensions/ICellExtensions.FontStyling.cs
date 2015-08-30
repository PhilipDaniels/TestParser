using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Extensions to the NPOI ICell class.
    /// </summary>
    public static partial class ICellExtensions
    {
        /// <summary>
        /// Sets the font weight.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="fontWeight">The font weight.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FontWeight(this ICell cell, FontBoldWeight fontWeight)
        {
            return new FluentCell(cell).FontWeight(fontWeight);
        }

        /// <summary>
        /// Sets the charset.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="charset">The charset.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Charset(this ICell cell, short charset)
        {
            return new FluentCell(cell).Charset(charset);
        }

        /// <summary>
        /// Sets the font color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="color">The color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Color(this ICell cell, short color)
        {
            return new FluentCell(cell).Color(color);
        }

        /// <summary>
        /// Sets the font height.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="fontHeight">Height of the font.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FontHeight(this ICell cell, double fontHeight)
        {
            return new FluentCell(cell).FontHeight(fontHeight);
        }

        /// <summary>
        /// Sets the font height in points.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="fontHeightInPoints">The font height in points.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FontHeightInPoints(this ICell cell, short fontHeightInPoints)
        {
            return new FluentCell(cell).FontHeightInPoints(fontHeightInPoints);
        }

        /// <summary>
        /// Sets the font name.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="fontName">Name of the font.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FontName(this ICell cell, string fontName)
        {
            return new FluentCell(cell).FontName(fontName);
        }

        /// <summary>
        /// Sets italics on or off.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="italic">if set to <c>true</c> [italic].</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Italic(this ICell cell, bool italic)
        {
            return new FluentCell(cell).Italic(italic);
        }

        /// <summary>
        /// Sets strikeout on or off.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="strikeout">if set to <c>true</c> [strikeout].</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Strikeout(this ICell cell, bool strikeout)
        {
            return new FluentCell(cell).Strikeout(strikeout);
        }

        /// <summary>
        /// Sets superscript type.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="superScript">The super script.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell SuperScript(this ICell cell, FontSuperScript superScript)
        {
            return new FluentCell(cell).SuperScript(superScript);
        }

        /// <summary>
        /// Sets the underline type.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="underline">The underline type.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Underline(this ICell cell, FontUnderlineType underline)
        {
            return new FluentCell(cell).Underline(underline);
        }
    }
}
