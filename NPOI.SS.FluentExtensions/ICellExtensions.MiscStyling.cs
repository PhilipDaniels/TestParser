using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Extensions to the NPOI ICell class.
    /// </summary>
    public static partial class ICellExtensions
    {
        /// <summary>
        /// Sets the format string.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="format">The format string.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Format(this ICell cell, string format)
        {
            return new FluentCell(cell).Format(format);
        }

        /// <summary>
        /// Sets the format string to long date format, <c>"yyyy-MM-dd HH:mm:ss"</c>.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FormatLongDate(this ICell cell)
        {
            return new FluentCell(cell).FormatLongDate();
        }

        /// <summary>
        /// Sets the style of all borders.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderStyle">The border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderAll(this ICell cell, BorderStyle borderStyle)
        {
            return new FluentCell(cell).BorderAll(borderStyle);
        }

        /// <summary>
        /// Sets the color of all borders.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderColorAll(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).BorderColorAll(colorIndex);
        }

        /// <summary>
        /// Sets a cell to fill with a particular color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell SolidFillColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).SolidFill(colorIndex);
        }
    }
}
