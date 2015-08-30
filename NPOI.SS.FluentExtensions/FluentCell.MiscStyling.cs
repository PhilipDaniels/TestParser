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
        /// Sets the format string.
        /// </summary>
        /// <param name="format">The format string.</param>
        /// <returns>The cell.</returns>
        public FluentCell Format(string format)
        {
            Style.Format = format;
            return this;
        }

        /// <summary>
        /// Sets the format string to long date format, <c>"yyyy-MM-dd HH:mm:ss"</c>.
        /// </summary>
        /// <returns>The cell.</returns>
        public FluentCell FormatLongDate()
        {
            return Format("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// Sets the style of all borders.
        /// </summary>
        /// <param name="borderStyle">The border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderAll(BorderStyle borderStyle)
        {
            Style.BorderTop = Style.BorderRight = Style.BorderBottom = Style.BorderLeft = borderStyle;
            return this;
        }

        /// <summary>
        /// Sets the color of all borders.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderColorAll(short colorIndex)
        {
            Style.TopBorderColor = Style.RightBorderColor = Style.BottomBorderColor = Style.LeftBorderColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets a cell to fill with a particular color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell SolidFill(short colorIndex)
        {
            Style.FillForegroundColor = colorIndex;
            Style.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            return this;
        }
    }
}
