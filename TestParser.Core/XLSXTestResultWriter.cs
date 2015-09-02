using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TestParser.Core
{
    public partial class XLSXTestResultWriter : ITestResultWriter
    {
        IWorkbook workbook;
        ParsedData testResults;
        readonly string yellowBandString;
        readonly string greenBandString;

        public XLSXTestResultWriter(int yellowBand, int greenBand)
        {
            yellowBandString = (((decimal)yellowBand) / 100m).ToString();
            greenBandString = (((decimal)greenBand) / 100m).ToString();
        }

        public void WriteResults(Stream s, ParsedData testResults)
        {
            workbook = new XSSFWorkbook();
            this.testResults = testResults;

            CreateResultsSummarySheet(workbook.CreateSheet("TestResults - Summary"));
            CreateResultsSheet(workbook.CreateSheet("TestResults"));
            if (testResults.CoverageData.Count() > 0)
            {
                CreateCoverageSummarySheet(workbook.CreateSheet("CoverageResults - Summary"));
                CreateCoverageSheet(workbook.CreateSheet("CoverageResults"));
            }

            workbook.Write(s);
        }

        IConditionalFormattingRule[] MakePercentageConditionalFormattingRules(ISheet sheet)
        {
            IConditionalFormattingRule rule1 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, greenBandString);
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
            fill1.FillPattern = (short)FillPattern.SolidForeground;

            IConditionalFormattingRule rule2 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, yellowBandString);
            IPatternFormatting fill2 = rule2.CreatePatternFormatting();
            fill2.FillBackgroundColor = IndexedColors.Yellow.Index;
            fill2.FillPattern = (short)FillPattern.SolidForeground;

            IConditionalFormattingRule rule3 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, yellowBandString);
            IPatternFormatting fill3 = rule3.CreatePatternFormatting();
            fill3.FillBackgroundColor = IndexedColors.Red.Index;
            fill3.FillPattern = (short)FillPattern.SolidForeground;

            var rules = new IConditionalFormattingRule[] { rule1, rule2, rule3 };
            return rules;
        }
    }
}
