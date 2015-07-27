using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace TestParser.Core.XL
{
    /// <summary>
    /// Contains higher-level "macro" styles, but still generally applicable to any application.
    /// </summary>
    public static class MacroPropertyStyling
    {
        public static FluentStyledCell SolidFillColor(this ICell cell, short colorIndex)
        {
            return cell.FillForegroundColor(colorIndex).FillPattern(FillPattern.SolidForeground);
        }

        public static FluentStyledCell SolidFillColor(this FluentStyledCell styledCell, short colorIndex)
        {
            return styledCell.FillForegroundColor(colorIndex).FillPattern(FillPattern.SolidForeground);
        }
    }
}
