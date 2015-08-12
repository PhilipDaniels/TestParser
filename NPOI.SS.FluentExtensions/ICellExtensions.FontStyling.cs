using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public static partial class ICellExtensions
    {
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
    }
}
