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
        /// Sets the horizontal alignment.
        /// </summary>
        /// <param name="alignment">The alignment.</param>
        /// <returns>The cell.</returns>
        public FluentCell Alignment(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }

        /// <summary>
        /// Sets the bottom border.
        /// </summary>
        /// <param name="borderBottom">The bottom border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderBottom(BorderStyle borderBottom)
        {
            Style.BorderBottom = borderBottom;
            return this;
        }

        /// <summary>
        /// Sets the diagonal border.
        /// </summary>
        /// <param name="borderDiagonal">The diagonal border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderDiagonal(BorderDiagonal borderDiagonal)
        {
            Style.BorderDiagonal = borderDiagonal;
            return this;
        }

        /// <summary>
        /// Sets the diagonal border color.
        /// </summary>
        /// <param name="borderDiagonalColor">Color of the diagonal border.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderDiagonalColor(short borderDiagonalColor)
        {
            Style.BorderDiagonalColor = borderDiagonalColor;
            return this;
        }

        /// <summary>
        /// Sets the diagonal border line style.
        /// </summary>
        /// <param name="borderDiagonalLineStyle">The diagonal border line style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderDiagonalLineStyle(BorderStyle borderDiagonalLineStyle)
        {
            Style.BorderDiagonalLineStyle = borderDiagonalLineStyle;
            return this;
        }

        /// <summary>
        /// Sets the left border.
        /// </summary>
        /// <param name="borderLeft">The left border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderLeft(BorderStyle borderLeft)
        {
            Style.BorderLeft = borderLeft;
            return this;
        }

        /// <summary>
        /// Sets the right border.
        /// </summary>
        /// <param name="borderRight">The right border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderRight(BorderStyle borderRight)
        {
            Style.BorderRight = borderRight;
            return this;
        }

        /// <summary>
        /// Sets the top border.
        /// </summary>
        /// <param name="borderTop">The top border style.</param>
        /// <returns>The cell.</returns>
        public FluentCell BorderTop(BorderStyle borderTop)
        {
            Style.BorderTop = borderTop;
            return this;
        }

        /// <summary>
        /// Sets the bottom border color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell BottomBorderColor(short colorIndex)
        {
            Style.BottomBorderColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the data format.
        /// </summary>
        /// <param name="dataFormat">The data format.</param>
        /// <returns>The cell.</returns>
        public FluentCell DataFormat(short dataFormat)
        {
            Style.DataFormat = dataFormat;
            return this;
        }

        /// <summary>
        /// Sets the background fill color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell FillBackgroundColor(short colorIndex)
        {
            Style.FillBackgroundColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the foreground fill color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell FillForegroundColor(short colorIndex)
        {
            Style.FillForegroundColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the fill pattern.
        /// </summary>
        /// <param name="fillPattern">The fill pattern.</param>
        /// <returns>The cell.</returns>
        public FluentCell FillPattern(FillPattern fillPattern)
        {
            Style.FillPattern = fillPattern;
            return this;
        }

        /// <summary>
        /// Sets the indention.
        /// </summary>
        /// <param name="indention">The indention.</param>
        /// <returns>The cell.</returns>
        public FluentCell Indention(short indention)
        {
            Style.Indention = indention;
            return this;
        }

        /// <summary>
        /// Sets the left border color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell LeftBorderColor(short colorIndex)
        {
            Style.LeftBorderColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the right border color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell RightBorderColor(short colorIndex)
        {
            Style.RightBorderColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the rotation.
        /// </summary>
        /// <param name="rotation">The rotation.</param>
        /// <returns>The cell.</returns>
        public FluentCell Rotation(short rotation)
        {
            Style.Rotation = rotation;
            return this;
        }

        /// <summary>
        /// Sets 'shrink-to-fit' on or off.
        /// </summary>
        /// <param name="shrinkToFit">if set to <c>true</c> then shrink text to fit.</param>
        /// <returns>The cell.</returns>
        public FluentCell ShrinkToFit(bool shrinkToFit)
        {
            Style.ShrinkToFit = shrinkToFit;
            return this;
        }

        /// <summary>
        /// Sets the top border color.
        /// </summary>
        /// <param name="colorIndex">Index of the color.</param>
        /// <returns>The cell.</returns>
        public FluentCell TopBorderColor(short colorIndex)
        {
            Style.TopBorderColor = colorIndex;
            return this;
        }

        /// <summary>
        /// Sets the vertical alignment.
        /// </summary>
        /// <param name="verticalAlignment">The vertical alignment.</param>
        /// <returns>The cell.</returns>
        public FluentCell VerticalAlignment(VerticalAlignment verticalAlignment)
        {
            Style.VerticalAlignment = verticalAlignment;
            return this;
        }

        /// <summary>
        /// Sets wrap text on or off.
        /// </summary>
        /// <param name="wrapText">if set to <c>true</c> [wrap text].</param>
        /// <returns>The cell.</returns>
        public FluentCell WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
    }
}
