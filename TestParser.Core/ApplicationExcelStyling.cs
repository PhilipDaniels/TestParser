using NPOI.HSSF.Util;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;

namespace TestParser.Core
{
    /// <summary>
    /// The highest level of property styling, where we establish styles
    /// that are specific to our application.
    /// </summary>
    public static class ApplicationExcelStyling
    {
        public static FluentCell HeaderStyle(this ICell cell)
        {
            return new FluentCell(cell).HeaderStyle();
        }

        public static FluentCell HeaderStyle(this FluentCell cell)
        {
            return cell.SolidFill(HSSFColor.Grey25Percent.Index)
                       .FontWeight(FontBoldWeight.Bold);
        }

        public static FluentCell LargeHeaderStyle(this ICell cell)
        {
            return new FluentCell(cell).LargeHeaderStyle();
        }

        public static FluentCell LargeHeaderStyle(this FluentCell cell)
        {
            return cell.SolidFill(HSSFColor.Grey25Percent.Index)
                       .FontWeight(FontBoldWeight.Bold)
                       .Italic(true)
                       .FontHeightInPoints(20);
        }

        public static FluentCell SummaryStyle(this ICell cell)
        {
            return new FluentCell(cell).SummaryStyle();
        }

        public static FluentCell SummaryStyle(this FluentCell cell)
        {
            return cell.SolidFill(HSSFColor.Grey25Percent.Index)
                       .FontWeight(FontBoldWeight.Bold)
                       .BorderAll(BorderStyle.Thin);
        }

        public static FluentCell FormatPercentage(this ICell cell)
        {
            return new FluentCell(cell).FormatPercentage();
        }

        public static FluentCell FormatPercentage(this FluentCell styledCell)
        {
            return styledCell.Format("0%");
        }
    }
}
