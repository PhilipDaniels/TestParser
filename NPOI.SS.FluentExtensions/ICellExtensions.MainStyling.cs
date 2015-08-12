using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public static partial class ICellExtensions
    {
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
    }
}
