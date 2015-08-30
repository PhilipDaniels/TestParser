using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Extensions to the NPOI ICell class.
    /// </summary>
    public static partial class ICellExtensions
    {
        /// <summary>
        /// Sets the horizontal alignment.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="alignment">The alignment.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Alignment(this ICell cell, HorizontalAlignment alignment)
        {
            return new FluentCell(cell).Alignment(alignment);
        }

        /// <summary>
        /// Sets the bottom border.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderBottom">The bottom border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderBottom(this ICell cell, BorderStyle borderBottom)
        {
            return new FluentCell(cell).BorderBottom(borderBottom);
        }

        /// <summary>
        /// Sets the diagonal border.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderDiagonal">The diagonal border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderDiagonal(this ICell cell, BorderDiagonal borderDiagonal)
        {
            return new FluentCell(cell).BorderDiagonal(borderDiagonal);
        }

        /// <summary>
        /// Sets the diagonal border color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderDiagonalColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).BorderDiagonalColor(colorIndex);
        }

        /// <summary>
        /// Sets the diagonal border line style.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderDiagonalLineStyle">The diagonal border line style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderDiagonalLineStyle(this ICell cell, BorderStyle borderDiagonalLineStyle)
        {
            return new FluentCell(cell).BorderDiagonalLineStyle(borderDiagonalLineStyle);
        }

        /// <summary>
        /// Sets the left border.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderLeft">The left border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderLeft(this ICell cell, BorderStyle borderLeft)
        {
            return new FluentCell(cell).BorderLeft(borderLeft);
        }

        /// <summary>
        /// Sets the right border.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderRight">The right border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderRight(this ICell cell, BorderStyle borderRight)
        {
            return new FluentCell(cell).BorderRight(borderRight);
        }

        /// <summary>
        /// Sets the top border.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="borderTop">The top border style.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BorderTop(this ICell cell, BorderStyle borderTop)
        {
            return new FluentCell(cell).BorderTop(borderTop);
        }

        /// <summary>
        /// Sets the bottom border color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell BottomBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).BottomBorderColor(colorIndex);
        }

        /// <summary>
        /// Sets the data format.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="dataFormat">The data format.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell DataFormat(this ICell cell, short dataFormat)
        {
            return new FluentCell(cell).DataFormat(dataFormat);
        }

        /// <summary>
        /// Sets the background fill color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FillBackgroundColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).FillBackgroundColor(colorIndex);
        }

        /// <summary>
        /// Sets the foreground fill color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FillForegroundColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).FillForegroundColor(colorIndex);
        }

        /// <summary>
        /// Sets the fill pattern.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="fillPattern">The fill pattern.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell FillPattern(this ICell cell, FillPattern fillPattern)
        {
            return new FluentCell(cell).FillPattern(fillPattern);
        }

        /// <summary>
        /// Sets the indention.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="Indention">The indention.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Indention(this ICell cell, short Indention)
        {
            return new FluentCell(cell).Indention(Indention);
        }

        /// <summary>
        /// Sets the left border color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell LeftBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).LeftBorderColor(colorIndex);
        }

        /// <summary>
        /// Sets the right border color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell RightBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).RightBorderColor(colorIndex);
        }

        /// <summary>
        /// Sets the rotation.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="rotation">The rotation.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell Rotation(this ICell cell, short rotation)
        {
            return new FluentCell(cell).Rotation(rotation);
        }

        /// <summary>
        /// Sets 'shrink-to-fit' on or off.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="shrinkToFit">if set to <c>true</c> then shrink text to fit.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell ShrinkToFit(this ICell cell, bool shrinkToFit)
        {
            return new FluentCell(cell).ShrinkToFit(shrinkToFit);
        }

        /// <summary>
        /// Sets the top border color.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell TopBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).TopBorderColor(colorIndex);
        }

        /// <summary>
        /// Sets the vertical alignment.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="verticalAlignment">The vertical alignment.</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell VerticalAlignment(this ICell cell, VerticalAlignment verticalAlignment)
        {
            return new FluentCell(cell).VerticalAlignment(verticalAlignment);
        }

        /// <summary>
        /// Sets wrap text on or off.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <param name="wrapText">if set to <c>true</c> [wrap text].</param>
        /// <returns>A <seealso cref="FluentCell"/>, for further styling.</returns>
        public static FluentCell WrapText(this ICell cell, bool wrapText)
        {
            return new FluentCell(cell).WrapText(wrapText);
        }
    }
}
