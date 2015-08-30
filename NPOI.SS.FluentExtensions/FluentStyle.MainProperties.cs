using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Provides a data structure which can be used to accumulate
    /// styling options via a Fluent API, until they need to be applied
    /// to a cell.
    /// </summary>
    public partial class FluentStyle
    {
        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        /// <value>
        /// The alignment.
        /// </value>
        public HorizontalAlignment? Alignment { get; set; }

        /// <summary>
        /// Gets or sets the border bottom style.
        /// </summary>
        /// <value>
        /// The border bottom.
        /// </value>
        public BorderStyle? BorderBottom { get; set; }

        /// <summary>
        /// Gets or sets the border diagonal.
        /// </summary>
        /// <value>
        /// The border diagonal.
        /// </value>
        public BorderDiagonal? BorderDiagonal { get; set; }

        /// <summary>
        /// Gets or sets the diagonal border color.
        /// </summary>
        /// <value>
        /// The color of the border diagonal.
        /// </value>
        public short? BorderDiagonalColor { get; set; }

        /// <summary>
        /// Gets or sets the diagonal border line style.
        /// </summary>
        /// <value>
        /// The border diagonal line style.
        /// </value>
        public BorderStyle? BorderDiagonalLineStyle { get; set; }

        /// <summary>
        /// Gets or sets the left border style.
        /// </summary>
        /// <value>
        /// The border left.
        /// </value>
        public BorderStyle? BorderLeft { get; set; }

        /// <summary>
        /// Gets or sets the right border style.
        /// </summary>
        /// <value>
        /// The border right.
        /// </value>
        public BorderStyle? BorderRight { get; set; }

        /// <summary>
        /// Gets or sets the top border style.
        /// </summary>
        /// <value>
        /// The border top.
        /// </value>
        public BorderStyle? BorderTop { get; set; }

        /// <summary>
        /// Gets or sets the color of the bottom border.
        /// </summary>
        /// <value>
        /// The color of the bottom border.
        /// </value>
        public short? BottomBorderColor { get; set; }

        /// <summary>
        /// Gets or sets the data format.
        /// </summary>
        /// <value>
        /// The data format.
        /// </value>
        public short? DataFormat { get; set; }

        /// <summary>
        /// Gets or sets the background fill color.
        /// </summary>
        /// <value>
        /// The color of the fill background.
        /// </value>
        public short? FillBackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the foreground fill color.
        /// </summary>
        /// <value>
        /// The color of the fill foreground.
        /// </value>
        public short? FillForegroundColor { get; set; }

        /// <summary>
        /// Gets or sets the fill pattern.
        /// </summary>
        /// <value>
        /// The fill pattern.
        /// </value>
        public FillPattern? FillPattern { get; set; }

        /// <summary>
        /// Gets or sets the indention.
        /// </summary>
        /// <value>
        /// The indention.
        /// </value>
        public short? Indention { get; set; }

        /// <summary>
        /// Gets or sets the color of the left border.
        /// </summary>
        /// <value>
        /// The color of the left border.
        /// </value>
        public short? LeftBorderColor { get; set; }

        /// <summary>
        /// Gets or sets the color of the right border.
        /// </summary>
        /// <value>
        /// The color of the right border.
        /// </value>
        public short? RightBorderColor { get; set; }

        /// <summary>
        /// Gets or sets the rotation.
        /// </summary>
        /// <value>
        /// The rotation.
        /// </value>
        public short? Rotation { get; set; }

        /// <summary>
        /// Gets or sets the shrink to fit setting.
        /// </summary>
        /// <value>
        /// The shrink to fit setting.
        /// </value>
        public bool? ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the color of the top border.
        /// </summary>
        /// <value>
        /// The color of the top border.
        /// </value>
        public short? TopBorderColor { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        /// <value>
        /// The vertical alignment.
        /// </value>
        public VerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the wrap text setting.
        /// </summary>
        /// <value>
        /// The wrap text setting.
        /// </value>
        public bool? WrapText { get; set; }

        /// <summary>
        /// Gets or sets the format, e.g. "0.00%".
        /// If used, overrides the <see cref="DataFormat"/> property when the style is applied.
        /// </summary>
        /// <value>
        /// The format.
        /// </value>
        public string Format { get; set; }
    }
}
