using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Bundle together an <c>ICell</c> and a <see cref="FluentStyle"/>
    /// so that they the cell can be configured with a fluent interface.
    /// </summary>
    public partial class FluentCell
    {
        static Dictionary<int, ICellStyle> cachedWorkbookStyles;

        static FluentCell()
        {
            cachedWorkbookStyles = new Dictionary<int, ICellStyle>();
        }

        /// <summary>
        /// Return the number of cached styles. Mainly for checking
        /// in the debugger that the caching is working as expected.
        /// </summary>
        public static int NumCachedStyles
        {
            get
            {
                return cachedWorkbookStyles.Count;
            }
        }

        /// <summary>
        /// Initialize a new instance of the <c>FluentCell</c>.
        /// </summary>
        /// <param name="cell">The NPOI cell this <c>FluentCell</c> is associated with.</param>
        public FluentCell(ICell cell)
        {
            Cell = cell;
            Style = new FluentStyle();
        }

        /// <summary>
        /// The NPOI cell this <c>FluentCell</c> is associated with.
        /// </summary>
        public ICell Cell { get; set; }

        /// <summary>
        /// The current FluentStyle. Calling <see cref="ApplyStyle"/> will apply it.
        /// </summary>
        public FluentStyle Style { get; set; }

        /// <summary>
        /// Applies the current style to the cell. Reuses a cached style
        /// if it exists (this is important because there is a limit to how
        /// many styles can be used in an Excel spreadsheet).
        /// </summary>
        /// <returns>The <c>FluentCell</c>.</returns>
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
    }
}
