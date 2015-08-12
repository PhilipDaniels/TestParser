using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public partial class FluentCell
    {
        public FluentCell FontWeight(FontBoldWeight fontWeight)
        {
            Style.FontWeight = fontWeight;
            return this;
        }

        public FluentCell Charset(short charset)
        {
            Style.Charset = charset;
            return this;
        }

        public FluentCell Color(short color)
        {
            Style.Color = color;
            return this;
        }

        public FluentCell FontHeight(double fontHeight)
        {
            Style.FontHeight = fontHeight;
            return this;
        }

        public FluentCell FontHeightInPoints(short fontHeightInPoints)
        {
            Style.FontHeightInPoints = fontHeightInPoints;
            return this;
        }

        public FluentCell FontName(string fontName)
        {
            Style.FontName = fontName;
            return this;
        }

        public FluentCell Italic(bool italic)
        {
            Style.Italic = italic;
            return this;
        }

        public FluentCell Strikeout(bool strikeout)
        {
            Style.Strikeout = strikeout;
            return this;
        }

        public FluentCell SuperScript(FontSuperScript superScript)
        {
            Style.SuperScript = superScript;
            return this;
        }

        public FluentCell Underline(FontUnderlineType underline)
        {
            Style.Underline = underline;
            return this;
        }
    }
}
