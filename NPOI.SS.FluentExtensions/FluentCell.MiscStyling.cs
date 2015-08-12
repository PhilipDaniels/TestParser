using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public partial class FluentCell
    {
        public FluentCell Format(string format)
        {
            Style.Format = format;
            return this;
        }

        public FluentCell FormatLongDate()
        {
            return Format("yyyy-MM-dd HH:mm:ss");
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
