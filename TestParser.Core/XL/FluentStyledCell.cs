using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace TestParser.Core.XL
{
    public class FluentStyledCell
    {
        static Dictionary<int, ICellStyle> cachedWorkbookStyles;

        static FluentStyledCell()
        {
            cachedWorkbookStyles = new Dictionary<int, ICellStyle>();
        }

        public ICell Cell { get; set; }
        public FluentStyle Style { get; set; }

        public FluentStyledCell ApplyStyle()
        {
            int styleHash = Style.GetHashCode();
            ICellStyle wbStyle;

            lock (cachedWorkbookStyles)
            {
                if (!cachedWorkbookStyles.TryGetValue(styleHash, out wbStyle))
                {
                    wbStyle = Cell.Sheet.Workbook.CreateCellStyle();
                    Style.ApplyStyle(wbStyle);
                    cachedWorkbookStyles.Add(styleHash, wbStyle);
                }
            }

            Cell.CellStyle = wbStyle;
            return this;
        }

        public FluentStyledCell Alignment(HorizontalAlignment alignment)
        {
            Style.Alignment = alignment;
            return this;
        }

        public FluentStyledCell BorderBottom(BorderStyle borderStyle)
        {
            Style.BorderBottom = borderStyle;
            return this;
        }

        public FluentStyledCell BorderDiagonal(BorderDiagonal borderDiagonal)
        {
            Style.BorderDiagonal = borderDiagonal;
            return this;
        }

        public FluentStyledCell BorderDiagonalColor(short borderDiagonalColor)
        {
            Style.BorderDiagonalColor = borderDiagonalColor;
            return this;
        }

        public FluentStyledCell BorderDiagonalLineStyle(BorderStyle borderDiagonalLineStyle)
        {
            Style.BorderDiagonalLineStyle = borderDiagonalLineStyle;
            return this;
        }

        public FluentStyledCell BorderLeft(BorderStyle borderLeft)
        {
            Style.BorderLeft = borderLeft;
            return this;
        }

        public FluentStyledCell BorderRight(BorderStyle borderRight)
        {
            Style.BorderRight = borderRight;
            return this;
        }

        public FluentStyledCell BorderTop(BorderStyle borderTop)
        {
            Style.BorderTop = borderTop;
            return this;
        }

        public FluentStyledCell BottomBorderColor(short colorIndex)
        {
            Style.BottomBorderColor = colorIndex;
            return this;
        }

        public FluentStyledCell DataFormat(short dataFormat)
        {
            Style.DataFormat = dataFormat;
            return this;
        }

        public FluentStyledCell FillBackgroundColor(short colorIndex)
        {
            Style.FillBackgroundColor = colorIndex;
            return this;
        }

        public FluentStyledCell FillForegroundColor(short colorIndex)
        {
            Style.FillForegroundColor = colorIndex;
            return this;
        }

        public FluentStyledCell FillPattern(FillPattern fillPattern)
        {
            Style.FillPattern = fillPattern;
            return this;
        }

        public FluentStyledCell Indention(short indention)
        {
            Style.Indention = indention;
            return this;
        }

        public FluentStyledCell LeftBorderColor(short colorIndex)
        {
            Style.LeftBorderColor = colorIndex;
            return this;
        }

        public FluentStyledCell RightBorderColor(short colorIndex)
        {
            Style.RightBorderColor = colorIndex;
            return this;
        }

        public FluentStyledCell Rotation(short rotation)
        {
            Style.Rotation = rotation;
            return this;
        }

        public FluentStyledCell ShrinkToFit(bool shrinkToFit)
        {
            Style.ShrinkToFit = shrinkToFit;
            return this;
        }

        public FluentStyledCell TopBorderColor(short colorIndex)
        {
            Style.TopBorderColor = colorIndex;
            return this;
        }

        public FluentStyledCell VerticalAlignment(VerticalAlignment verticalAlignment)
        {
            Style.VerticalAlignment = verticalAlignment;
            return this;
        }

        public FluentStyledCell WrapText(bool wrapText)
        {
            Style.WrapText = wrapText;
            return this;
        }
    }
}
