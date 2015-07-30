using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public class FluentStyle
    {
        #region Main cell style properties
        public HorizontalAlignment? Alignment { get; set; }
        public BorderStyle? BorderBottom { get; set; }
        public BorderDiagonal? BorderDiagonal { get; set; }
        public short? BorderDiagonalColor { get; set; }
        public BorderStyle? BorderDiagonalLineStyle { get; set; }
        public BorderStyle? BorderLeft { get; set; }
        public BorderStyle? BorderRight { get; set; }
        public BorderStyle? BorderTop { get; set; }
        public short? BottomBorderColor { get; set; }
        public short? DataFormat { get; set; }
        public short? FillBackgroundColor { get; set; }
        public short? FillForegroundColor { get; set; }
        public FillPattern? FillPattern { get; set; }
        public short? Indention { get; set; }
        public short? LeftBorderColor { get; set; }
        public short? RightBorderColor { get; set; }
        public short? Rotation { get; set; }
        public bool? ShrinkToFit { get; set; }
        public short? TopBorderColor { get; set; }
        public VerticalAlignment? VerticalAlignment { get; set; }
        public bool? WrapText { get; set; }
        #endregion

        #region Font properties
        public FontBoldWeight? FontWeight { get; set; }
        public short? Charset { get; set; }
        public short? Color { get; set; }
        public double? FontHeight { get; set; }
        public short? FontHeightInPoints { get; set; }
        public string FontName { get; set; }
        public bool? Italic { get; set; }
        public bool? Strikeout { get; set; }
        public FontSuperScript? SuperScript { get; set; }
        public FontUnderlineType? Underline { get; set; }
        #endregion

        /// <summary>
        /// Gets or sets the format, e.g. "0.00%".
        /// If used, overrides the <see cref="DataFormat"/> property when the style is applied.
        /// </summary>
        /// <value>
        /// The format.
        /// </value>
        public string Format { get; set; }

        public void ApplyStyle(IWorkbook workbook, ICellStyle destination)
        {
            // If users sets format string this overrides the DataFormat property.
            if (Format != null)
            {
                var dataFormat = workbook.CreateDataFormat();
                DataFormat = dataFormat.GetFormat(Format);
            }

            if (Alignment != null) destination.Alignment = Alignment.Value;
            if (BorderBottom != null) destination.BorderBottom = BorderBottom.Value;
            if (BorderDiagonal != null) destination.BorderDiagonal = BorderDiagonal.Value;
            if (BorderDiagonalColor != null) destination.BorderDiagonalColor = BorderDiagonalColor.Value;
            if (BorderDiagonalLineStyle != null) destination.BorderDiagonalLineStyle = BorderDiagonalLineStyle.Value;
            if (BorderLeft != null) destination.BorderLeft = BorderLeft.Value;
            if (BorderRight != null) destination.BorderRight = BorderRight.Value;
            if (BorderTop != null) destination.BorderTop = BorderTop.Value;
            if (BottomBorderColor != null) destination.BottomBorderColor = BottomBorderColor.Value;
            if (DataFormat != null) destination.DataFormat = DataFormat.Value;
            if (FillBackgroundColor != null) destination.FillBackgroundColor = FillBackgroundColor.Value;
            if (FillForegroundColor != null) destination.FillForegroundColor = FillForegroundColor.Value;
            if (FillPattern != null) destination.FillPattern = FillPattern.Value;
            if (Indention != null) destination.Indention = Indention.Value;
            if (LeftBorderColor != null) destination.LeftBorderColor = LeftBorderColor.Value;
            if (RightBorderColor != null) destination.RightBorderColor = RightBorderColor.Value;
            if (Rotation != null) destination.Rotation = Rotation.Value;
            if (ShrinkToFit != null) destination.ShrinkToFit = ShrinkToFit.Value;
            if (TopBorderColor != null) destination.TopBorderColor = TopBorderColor.Value;
            if (VerticalAlignment != null) destination.VerticalAlignment = VerticalAlignment.Value;
            if (WrapText != null) destination.WrapText = WrapText.Value;

            if (FontIsNeeded)
            {
                var font = workbook.CreateFont();

                if (FontWeight != null) font.Boldweight = (short)FontWeight.Value;
                if (Charset != null) font.Charset = Charset.Value;
                if (Color != null) font.Color = Color.Value;
                if (FontHeight != null) font.FontHeight = FontHeight.Value;
                if (FontHeightInPoints != null) font.FontHeightInPoints = FontHeightInPoints.Value;
                if (FontName != null) font.FontName = FontName;
                if (Italic != null) font.IsItalic = Italic.Value;
                if (Strikeout != null) font.IsStrikeout = Strikeout.Value;
                if (SuperScript != null) font.TypeOffset = SuperScript.Value;
                if (Underline != null) font.Underline = Underline.Value;

                destination.SetFont(font);
            }
        }

        /// <summary>
        /// Returns the hash code of the object.
        /// </summary>
        /// <returns>Hash code.</returns>
        public override int GetHashCode()
        {
            // See Jon Skeet: http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode/263416#263416
            // Overflow is fine, just wrap.
            // If writing this seems hard, try: https://bitbucket.org/JonHanna/spookilysharp/src, but check the licence.
            // Warning: GetHashCode *should* be implemented only on immutable fields to guard against the
            // case where you mutate an object that has been used as a key in a collection. In practice,
            // you probably don't need to worry about that since it is a dumb thing to do.
            unchecked
            {
                int hash = 17;
                if (Alignment != null) hash += 23 * Alignment.Value.GetHashCode();
                if (BorderBottom != null) hash += 23 * BorderBottom.Value.GetHashCode();
                if (BorderDiagonal != null) hash += 23 * BorderDiagonal.Value.GetHashCode();
                if (BorderDiagonalColor != null) hash += 23 * BorderDiagonalColor.Value.GetHashCode();
                if (BorderDiagonalLineStyle != null) hash += 23 * BorderDiagonalLineStyle.Value.GetHashCode();
                if (BorderLeft != null) hash += 23 * BorderLeft.Value.GetHashCode();
                if (BorderRight != null) hash += 23 * BorderRight.Value.GetHashCode();
                if (BorderTop != null) hash += 23 * BorderTop.Value.GetHashCode();
                if (BottomBorderColor != null) hash += 23 * BottomBorderColor.Value.GetHashCode();
                if (DataFormat != null) hash += 23 * DataFormat.Value.GetHashCode();
                if (FillBackgroundColor != null) hash += 23 * FillBackgroundColor.Value.GetHashCode();
                if (FillForegroundColor != null) hash += 23 * FillForegroundColor.Value.GetHashCode();
                if (FillPattern != null) hash += 23 * FillPattern.Value.GetHashCode();
                if (Indention != null) hash += 23 * Indention.Value.GetHashCode();
                if (LeftBorderColor != null) hash += 23 * LeftBorderColor.Value.GetHashCode();
                if (RightBorderColor != null) hash += 23 * RightBorderColor.Value.GetHashCode();
                if (Rotation != null) hash += 23 * Rotation.Value.GetHashCode();
                if (ShrinkToFit != null) hash += 23 * ShrinkToFit.Value.GetHashCode();
                if (TopBorderColor != null) hash += 23 * TopBorderColor.Value.GetHashCode();
                if (VerticalAlignment != null) hash += 23 * VerticalAlignment.Value.GetHashCode();
                if (WrapText != null) hash += 23 * WrapText.Value.GetHashCode();

                if (FontWeight != null) hash += 23 * FontWeight.Value.GetHashCode();
                if (Charset != null) hash += 23 * Charset.Value.GetHashCode();
                if (Color != null) hash += 23 * Color.Value.GetHashCode();
                if (FontHeight != null) hash += 23 * FontHeight.Value.GetHashCode();
                if (FontHeightInPoints != null) hash += 23 * FontHeightInPoints.Value.GetHashCode();
                if (FontName != null) hash += 23 * FontName.GetHashCode();
                if (Italic != null) hash += 23 * Italic.Value.GetHashCode();
                if (Strikeout != null) hash += 23 * Strikeout.Value.GetHashCode();
                if (SuperScript != null) hash += 23 * SuperScript.Value.GetHashCode();
                if (Underline != null) hash += 23 * Underline.Value.GetHashCode();

                if (Format != null) hash += 23 * Format.GetHashCode();

                return hash;
            }
        }

        /// <summary>
        /// Returns true if <paramref name="value"/> is equal to this StyleDTO.
        /// </summary>
        /// <param name="value">The object to check for equality.</param>
        /// <returns>True if the object is equal to this one.</returns>
        public override bool Equals(object value)
        {
            return Equals((FluentStyle)value);
        }

        /// <summary>
        /// Returns true if <paramref name="value"/> is equal to this StyleDTO.
        /// </summary>
        /// <param name="value">The object to check for equality.</param>
        /// <returns>True if the object is equal to this one.</returns>
        public bool Equals(FluentStyle value)
        {
            // Use object.ReferenceEquals to avoid infinite loops.
            if (object.ReferenceEquals(value, null))
                return false;

            // First check for an exact type match.
            if (!object.ReferenceEquals(GetType(), value.GetType()))
                return false;

            return Alignment == value.Alignment &&
                BorderBottom == value.BorderBottom &&
                BorderDiagonal == value.BorderDiagonal &&
                BorderDiagonalColor == value.BorderDiagonalColor &&
                BorderDiagonalLineStyle == value.BorderDiagonalLineStyle &&
                BorderLeft == value.BorderLeft &&
                BorderRight == value.BorderRight &&
                BorderTop == value.BorderTop &&
                BottomBorderColor == value.BottomBorderColor &&
                DataFormat == value.DataFormat &&
                FillBackgroundColor == value.FillBackgroundColor &&
                FillForegroundColor == value.FillForegroundColor &&
                FillPattern == value.FillPattern &&
                Indention == value.Indention &&
                LeftBorderColor == value.LeftBorderColor &&
                RightBorderColor == value.RightBorderColor &&
                Rotation == value.Rotation &&
                ShrinkToFit == value.ShrinkToFit &&
                TopBorderColor == value.TopBorderColor &&
                VerticalAlignment == value.VerticalAlignment &&
                WrapText == value.WrapText &&

                FontWeight == value.FontWeight &&
                Charset == value.Charset &&
                Color == value.Color &&
                FontHeight == value.FontHeight &&
                FontHeightInPoints == value.FontHeightInPoints &&
                FontName == value.FontName &&
                Italic == value.Italic &&
                Strikeout == value.Strikeout &&
                SuperScript == value.SuperScript &&
                Underline == value.Underline &&

                Format == value.Format;
        }

        /// <summary>
        /// Returns true if the two StyleDTO instances are equal.
        /// </summary>
        /// <param name="first">First object.</param>
        /// <param name="second">Second object.</param>
        /// <returns>True if the two objects are equal.</returns>
        public static bool operator ==(FluentStyle first, FluentStyle second)
        {
            // Use object.ReferenceEquals to avoid infinite loops.
            if (object.ReferenceEquals(first, null))
                return object.ReferenceEquals(second, null);
            else
                return first.Equals(second);
        }

        /// <summary>
        /// Returns true if the two StyleDTO instances are not equal.
        /// </summary>
        /// <param name="first">First object.</param>
        /// <param name="second">Second object.</param>
        /// <returns>True if the two objects are equal.</returns>
        public static bool operator !=(FluentStyle first, FluentStyle second)
        {
            return !(first == second);
        }

        bool FontIsNeeded
        {
            get
            {
                return FontWeight != null ||
                        Charset != null ||
                        Color != null ||
                        FontHeight != null ||
                        FontHeightInPoints != null ||
                        FontName != null ||
                        Italic != null ||
                        Strikeout != null ||
                        SuperScript != null ||
                        Underline != null;
            }
        }
    }
}
