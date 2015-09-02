using System;
using BassUtils;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public abstract class XLSXSheetWriterBase
    {
        protected readonly ISheet sheet;

        public int YellowBand { get; private set; }
        public int GreenBand { get; private set; }
        public string YellowBandString { get; private set; }
        public string GreenBandString { get; private set; }
        public IConditionalFormattingRule[] PercentageFormattingRules { get; private set; }

        protected XLSXSheetWriterBase(ISheet sheet)
        {
            this.sheet = sheet.ThrowIfNull("sheet");
        }

        public void SetPercentageBands(int yellowBand, int greenBand)
        {
            YellowBand = yellowBand;
            GreenBand = greenBand;
            YellowBandString = (((decimal)yellowBand) / 100m).ToString();
            GreenBandString = (((decimal)greenBand) / 100m).ToString();

            MakePercentageConditionalFormattingRules();
        }

        protected void ApplyPercentageFormatting(int column, int rowFromInclusive, int rowtoInclusive)
        {
            string colString = CellReference.ConvertNumToColString(column);
            string range = String.Format("{0}{1}:{0}{2}", colString, rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, PercentageFormattingRules);
        }

        protected void ApplyPercentageFormatting(int columnFromInclusive, int columnToInclusive, int rowFromInclusive, int rowtoInclusive)
        {
            string colStringFrom = CellReference.ConvertNumToColString(columnFromInclusive);
            string colStringTo = CellReference.ConvertNumToColString(columnToInclusive);
            string range = String.Format("{0}{1}:{2}{3}", colStringFrom, rowFromInclusive, colStringTo, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, PercentageFormattingRules);
        }

        void MakePercentageConditionalFormattingRules()
        {
            IConditionalFormattingRule rule1 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, GreenBandString);
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
            fill1.FillPattern = (short)FillPattern.SolidForeground;

            IConditionalFormattingRule rule2 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, YellowBandString);
            IPatternFormatting fill2 = rule2.CreatePatternFormatting();
            fill2.FillBackgroundColor = IndexedColors.Yellow.Index;
            fill2.FillPattern = (short)FillPattern.SolidForeground;

            IConditionalFormattingRule rule3 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, YellowBandString);
            IPatternFormatting fill3 = rule3.CreatePatternFormatting();
            fill3.FillBackgroundColor = IndexedColors.Red.Index;
            fill3.FillPattern = (short)FillPattern.SolidForeground;

            PercentageFormattingRules = new IConditionalFormattingRule[] { rule1, rule2, rule3 };
        }
    }
}
