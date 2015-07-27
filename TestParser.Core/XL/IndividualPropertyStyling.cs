using NPOI.SS.UserModel;

namespace TestParser.Core.XL
{
    /// <summary>
    /// Contains extension methods which map directly to adjusting one of the
    /// ICellStyle properties.
    /// </summary>
    public static class IndividualPropertyStyling
    {
        public static FluentStyledCell Alignment(this ICell cell, HorizontalAlignment alignment)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    Alignment = alignment
                }
            };
        }

        public static FluentStyledCell BorderBottom(this ICell cell, BorderStyle borderBottom)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderBottom = borderBottom
                }
            };
        }

        public static FluentStyledCell BorderDiagonal(this ICell cell, BorderDiagonal borderDiagonal)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderDiagonal = borderDiagonal
                }
            };
        }

        public static FluentStyledCell BorderDiagonalColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderDiagonalColor = colorIndex
                }
            };
        }

        public static FluentStyledCell BorderDiagonalLineStyle(this ICell cell, BorderStyle borderDiagonalLineStyle)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderDiagonalLineStyle = borderDiagonalLineStyle
                }
            };
        }

        public static FluentStyledCell BorderLeft(this ICell cell, BorderStyle borderLeft)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderLeft = borderLeft
                }
            };
        }

        public static FluentStyledCell BorderRight(this ICell cell, BorderStyle borderRight)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderRight = borderRight
                }
            };
        }

        public static FluentStyledCell BorderTop(this ICell cell, BorderStyle borderTop)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BorderTop = borderTop
                }
            };
        }

        public static FluentStyledCell BottomBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    BottomBorderColor = colorIndex
                }
            };
        }

        public static FluentStyledCell DataFormat(this ICell cell, short dataFormat)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    DataFormat = dataFormat
                }
            };
        }

        public static FluentStyledCell FillBackgroundColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    FillBackgroundColor = colorIndex
                }
            };
        }

        public static FluentStyledCell FillForegroundColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    FillForegroundColor = colorIndex
                }
            };
        }

        public static FluentStyledCell FillPattern(this ICell cell, FillPattern fillPattern)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    FillPattern = fillPattern
                }
            };
        }

        public static FluentStyledCell Indention(this ICell cell, short indention)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    Indention = indention
                }
            };
        }

        public static FluentStyledCell LeftBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    LeftBorderColor = colorIndex
                }
            };
        }

        public static FluentStyledCell RightBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    RightBorderColor = colorIndex
                }
            };
        }

        public static FluentStyledCell Rotation(this ICell cell, short rotation)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    Rotation = rotation
                }
            };
        }

        public static FluentStyledCell ShrinkToFit(this ICell cell, bool shrinkToFit)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    ShrinkToFit = shrinkToFit
                }
            };
        }

        public static FluentStyledCell TopBorderColor(this ICell cell, short colorIndex)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    TopBorderColor = colorIndex
                }
            };
        }

        public static FluentStyledCell VerticalAlignment(this ICell cell, VerticalAlignment verticalAlignment)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    VerticalAlignment = verticalAlignment
                }
            };
        }

        public static FluentStyledCell WrapText(this ICell cell, bool wrapText)
        {
            return new FluentStyledCell()
            {
                Cell = cell,
                Style = new FluentStyleDTO()
                {
                    WrapText = wrapText
                }
            };
        }
    }
}
