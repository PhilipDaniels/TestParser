using System.Collections.Generic;
using NPOI.HSSF.Util;
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
        public FluentStyleDTO Style { get; set; }

        public FluentStyledCell ApplyStyle()
        {
            int styleHash = Style.GetHashCode();
            ICellStyle wbStyle;

            lock (cachedWorkbookStyles)
            {
                if (!cachedWorkbookStyles.TryGetValue(styleHash, out wbStyle))
                {
                    wbStyle = Cell.Sheet.Workbook.CreateCellStyle();
                    Style.ApplyStyle(Cell.Sheet.Workbook, wbStyle);
                    cachedWorkbookStyles.Add(styleHash, wbStyle);
                }
            }

            Cell.CellStyle = wbStyle;
            return this;
        }

        #region Main cell style properties
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
        #endregion

        #region Font Properties
        public FluentStyledCell FontWeight(FontBoldWeight fontWeight)
        {
            Style.FontWeight = fontWeight;
            return this;
        }

        public FluentStyledCell Charset(short charset)
        {
            Style.Charset = charset;
            return this;
        }

        public FluentStyledCell Color(short color)
        {
            Style.Color = color;
            return this;
        }

        public FluentStyledCell FontHeight(double fontHeight)
        {
            Style.FontHeight = fontHeight;
            return this;
        }

        public FluentStyledCell FontHeightInPoints(short fontHeightInPoints)
        {
            Style.FontHeightInPoints = fontHeightInPoints;
            return this;
        }

        public FluentStyledCell FontName(string fontName)
        {
            Style.FontName = fontName;
            return this;
        }

        public FluentStyledCell Italic(bool italic)
        {
            Style.Italic = italic;
            return this;
        }

        public FluentStyledCell Strikeout(bool strikeout)
        {
            Style.Strikeout = strikeout;
            return this;
        }

        public FluentStyledCell SuperScript(FontSuperScript superScript)
        {
            Style.SuperScript = superScript;
            return this;
        }

        public FluentStyledCell Underline(FontUnderlineType underline)
        {
            Style.Underline = underline;
            return this;
        }
        #endregion

        public FluentStyledCell Format(string format)
        {
            Style.Format = format;
            return this;
        }
    }
}
