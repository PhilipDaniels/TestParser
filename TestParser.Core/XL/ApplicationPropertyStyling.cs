using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace TestParser.Core.XL
{
    /// <summary>
    /// The highest level of property styling, where we establish styles
    /// that are specific to our application.
    /// </summary>
    public static class ApplicationPropertyStyling
    {
        public static FluentStyledCell HeaderStyle(this ICell cell)
        {
            var fs = new FluentStyledCell() { Cell = cell, Style = new FluentStyleDTO() };
            fs.HeaderStyle();
            return fs;
        }

        public static FluentStyledCell HeaderStyle(this FluentStyledCell styledCell)
        {
            return styledCell.SolidFillColor(HSSFColor.Grey25Percent.Index).FontWeight(FontBoldWeight.Bold);
        }

        public static FluentStyledCell LargeHeaderStyle(this ICell cell)
        {
            var fs = new FluentStyledCell() { Cell = cell, Style = new FluentStyleDTO() };
            fs.LargeHeaderStyle();
            return fs;
        }

        public static FluentStyledCell LargeHeaderStyle(this FluentStyledCell styledCell)
        {
            return styledCell.SolidFillColor(HSSFColor.Grey25Percent.Index).FontWeight(FontBoldWeight.Bold)
                .Italic(true).FontHeightInPoints(20);
        }

        public static FluentStyledCell FormatPercentage(this ICell cell)
        {
            var fs = new FluentStyledCell() { Cell = cell, Style = new FluentStyleDTO() };
            fs.FormatPercentage();
            return fs;
        }

        public static FluentStyledCell FormatPercentage(this FluentStyledCell styledCell)
        {
            return styledCell.Format("0.00%");
        }
    }
}
