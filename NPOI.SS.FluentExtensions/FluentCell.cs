using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public class FluentCell
    {
        static Dictionary<int, ICellStyle> cachedWorkbookStyles;

        static FluentCell()
        {
            cachedWorkbookStyles = new Dictionary<int, ICellStyle>();
        }

        public static int NumCachedStyles
        {
            get
            {
                return cachedWorkbookStyles.Count;
            }
        }

        public FluentCell(ICell cell)
        {
            Cell = cell;
            Style = new FluentStyle();
        }

        public ICell Cell { get; set; }
        public FluentStyle Style { get; set; }

        public FluentCell ApplyStyle()
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
        #endregion

        #region Font Properties
        public FluentCell FontWeight(FontBoldWeight fontWeight)
        {
            Style.FontWeight = fontWeight;
            return this;
        }

        public FluentCell Charset(short charset)
        {
            Style.Charset = charset;
            return this;
        }

        public FluentCell Color(short color)
        {
            Style.Color = color;
            return this;
        }

        public FluentCell FontHeight(double fontHeight)
        {
            Style.FontHeight = fontHeight;
            return this;
        }

        public FluentCell FontHeightInPoints(short fontHeightInPoints)
        {
            Style.FontHeightInPoints = fontHeightInPoints;
            return this;
        }

        public FluentCell FontName(string fontName)
        {
            Style.FontName = fontName;
            return this;
        }

        public FluentCell Italic(bool italic)
        {
            Style.Italic = italic;
            return this;
        }

        public FluentCell Strikeout(bool strikeout)
        {
            Style.Strikeout = strikeout;
            return this;
        }

        public FluentCell SuperScript(FontSuperScript superScript)
        {
            Style.SuperScript = superScript;
            return this;
        }

        public FluentCell Underline(FontUnderlineType underline)
        {
            Style.Underline = underline;
            return this;
        }
        #endregion

        public FluentCell Format(string format)
        {
            Style.Format = format;
            return this;
        }

        public FluentCell FormatLongDate()
        {
            Format("yyyy-MM-dd HH:mm:ss");
            return this;
        }

        public FluentCell BorderAll(BorderStyle borderStyle)
        {
            Style.BorderTop = Style.BorderRight = Style.BorderBottom = Style.BorderLeft = borderStyle;
            return this;
        }

        public FluentCell SolidFill(short colorIndex)
        {
            Style.FillForegroundColor = colorIndex;
            Style.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            return this;
        }
    }
}
