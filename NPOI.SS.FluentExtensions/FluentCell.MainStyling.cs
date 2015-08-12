using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public partial class FluentCell
    {
        public FluentCell Alignment(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }

        public FluentCell BorderBottom(BorderStyle borderStyle)
        {
            Style.BorderBottom = borderStyle;
            return this;
        }

        public FluentCell BorderDiagonal(BorderDiagonal borderDiagonal)
        {
            Style.BorderDiagonal = borderDiagonal;
            return this;
        }

        public FluentCell BorderDiagonalColor(short borderDiagonalColor)
        {
            Style.BorderDiagonalColor = borderDiagonalColor;
            return this;
        }

        public FluentCell BorderDiagonalLineStyle(BorderStyle borderDiagonalLineStyle)
        {
            Style.BorderDiagonalLineStyle = borderDiagonalLineStyle;
            return this;
        }

        public FluentCell BorderLeft(BorderStyle borderLeft)
        {
            Style.BorderLeft = borderLeft;
            return this;
        }

        public FluentCell BorderRight(BorderStyle borderRight)
        {
            Style.BorderRight = borderRight;
            return this;
        }

        public FluentCell BorderTop(BorderStyle borderTop)
        {
            Style.BorderTop = borderTop;
            return this;
        }

        public FluentCell BottomBorderColor(short colorIndex)
        {
            Style.BottomBorderColor = colorIndex;
            return this;
        }

        public FluentCell DataFormat(short dataFormat)
        {
            Style.DataFormat = dataFormat;
            return this;
        }

        public FluentCell FillBackgroundColor(short colorIndex)
        {
            Style.FillBackgroundColor = colorIndex;
            return this;
        }

        public FluentCell FillForegroundColor(short colorIndex)
        {
            Style.FillForegroundColor = colorIndex;
            return this;
        }

        public FluentCell FillPattern(FillPattern fillPattern)
        {
            Style.FillPattern = fillPattern;
            return this;
        }

        public FluentCell Indention(short indention)
        {
            Style.Indention = indention;
            return this;
        }

        public FluentCell LeftBorderColor(short colorIndex)
        {
            Style.LeftBorderColor = colorIndex;
            return this;
        }

        public FluentCell RightBorderColor(short colorIndex)
        {
            Style.RightBorderColor = colorIndex;
            return this;
        }

        public FluentCell Rotation(short rotation)
        {
            Style.Rotation = rotation;
            return this;
        }

        public FluentCell ShrinkToFit(bool shrinkToFit)
        {
            Style.ShrinkToFit = shrinkToFit;
            return this;
        }

        public FluentCell TopBorderColor(short colorIndex)
        {
            Style.TopBorderColor = colorIndex;
            return this;
        }

        public FluentCell VerticalAlignment(VerticalAlignment verticalAlignment)
        {
            Style.VerticalAlignment = verticalAlignment;
            return this;
        }

        public FluentCell WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
    }
}
