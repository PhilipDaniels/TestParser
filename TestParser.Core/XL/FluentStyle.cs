using NPOI.SS.UserModel;

namespace TestParser.Core.XL
{
    public class FluentStyle
    {
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

        public void ApplyStyle(ICellStyle destination)
        {
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
                WrapText == value.WrapText;
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
    }
}
