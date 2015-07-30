using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public static class ICellExtensions
    {
        #region Main cell style properties
        public static FluentCell Alignment(this ICell cell, HorizontalAlignment alignment)
        {
            return new FluentCell(cell).Alignment(alignment);
        }

        public static FluentCell BorderBottom(this ICell cell, BorderStyle borderBottom)
        {
            return new FluentCell(cell).BorderBottom(borderBottom);
        }

        public static FluentCell BorderDiagonal(this ICell cell, BorderDiagonal borderDiagonal)
        {
            return new FluentCell(cell).BorderDiagonal(borderDiagonal);
        }

        public static FluentCell BorderDiagonalColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).BorderDiagonalColor(colorIndex);
        }

        public static FluentCell BorderDiagonalLineStyle(this ICell cell, BorderStyle borderDiagonalLineStyle)
        {
            return new FluentCell(cell).BorderDiagonalLineStyle(borderDiagonalLineStyle);
        }

        public static FluentCell BorderLeft(this ICell cell, BorderStyle borderLeft)
        {
            return new FluentCell(cell).BorderLeft(borderLeft);
        }

        public static FluentCell BorderRight(this ICell cell, BorderStyle borderRight)
        {
            return new FluentCell(cell).BorderRight(borderRight);
        }

        public static FluentCell BorderTop(this ICell cell, BorderStyle borderTop)
        {
            return new FluentCell(cell).BorderTop(borderTop);
        }

        public static FluentCell BottomBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).BottomBorderColor(colorIndex);
        }

        public static FluentCell DataFormat(this ICell cell, short dataFormat)
        {
            return new FluentCell(cell).DataFormat(dataFormat);
        }

        public static FluentCell FillBackgroundColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).FillBackgroundColor(colorIndex);
        }

        public static FluentCell FillForegroundColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).FillForegroundColor(colorIndex);
        }

        public static FluentCell FillPattern(this ICell cell, FillPattern fillPattern)
        {
            return new FluentCell(cell).FillPattern(fillPattern);
        }

        public static FluentCell Indention(this ICell cell, short Indention)
        {
            return new FluentCell(cell).Indention(Indention);
        }

        public static FluentCell LeftBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).LeftBorderColor(colorIndex);
        }

        public static FluentCell RightBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).RightBorderColor(colorIndex);
        }

        public static FluentCell Rotation(this ICell cell, short rotation)
        {
            return new FluentCell(cell).Rotation(rotation);
        }

        public static FluentCell ShrinkToFit(this ICell cell, bool shrinkToFit)
        {
            return new FluentCell(cell).ShrinkToFit(shrinkToFit);
        }

        public static FluentCell TopBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).TopBorderColor(colorIndex);
        }

        public static FluentCell VerticalAlignment(this ICell cell, VerticalAlignment verticalAlignment)
        {
            return new FluentCell(cell).VerticalAlignment(verticalAlignment);
        }

        public static FluentCell WrapText(this ICell cell, bool wrapText)
        {
            return new FluentCell(cell).WrapText(wrapText);
        }
        #endregion

        #region Font Properties
        public static FluentCell FontWeight(this ICell cell, FontBoldWeight fontWeight)
        {
            return new FluentCell(cell).FontWeight(fontWeight);
        }

        public static FluentCell Charset(this ICell cell, short charset)
        {
            return new FluentCell(cell).Charset(charset);
        }

        public static FluentCell Color(this ICell cell, short color)
        {
            return new FluentCell(cell).Color(color);
        }

        public static FluentCell FontHeight(this ICell cell, double fontHeight)
        {
            return new FluentCell(cell).FontHeight(fontHeight);
        }

        public static FluentCell FontHeightInPoints(this ICell cell, short fontHeightInPoints)
        {
            return new FluentCell(cell).FontHeightInPoints(fontHeightInPoints);
        }

        public static FluentCell FontName(this ICell cell, string fontName)
        {
            return new FluentCell(cell).FontName(fontName);
        }

        public static FluentCell Italic(this ICell cell, bool italic)
        {
            return new FluentCell(cell).Italic(italic);
        }

        public static FluentCell Strikeout(this ICell cell, bool strikeout)
        {
            return new FluentCell(cell).Strikeout(strikeout);
        }

        public static FluentCell SuperScript(this ICell cell, FontSuperScript superScript)
        {
            return new FluentCell(cell).SuperScript(superScript);
        }

        public static FluentCell Underline(this ICell cell, FontUnderlineType underline)
        {
            return new FluentCell(cell).Underline(underline);
        }
        #endregion

        public static FluentCell Format(this ICell cell, string format)
        {
            return new FluentCell(cell).Format(format);
        }

        public static FluentCell FormatLongDate(this ICell cell)
        {
            return new FluentCell(cell).FormatLongDate();
        }

        public static FluentCell BorderAll(this ICell cell, BorderStyle borderStyle)
        {
            return new FluentCell(cell).BorderAll(borderStyle);
        }

        public static FluentCell SolidFillColor(this ICell cell, short colorIndex)
        {
            return new FluentCell(cell).SolidFill(colorIndex);
        }
    }
}
