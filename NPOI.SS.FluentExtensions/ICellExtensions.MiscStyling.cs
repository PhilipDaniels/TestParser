using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public static partial class ICellExtensions
    {
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
