using System;
using System.Collections.Generic;
using System.Linq;
using BassUtils;
using NPOI.HSSF.Util;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public class XLSXTestResultSummarySheetWriter : XLSXSheetWriterBase
    {
        const string ByTestAssembly = "Summary By Test Assembly";
        const string ByTestClass = "Summary By Test Class";

        const int ColAssembly = 0;
        const int ColClass = 1;
        const int ColTime = 2;
        const int ColTimeHuman = 3;
        const int ColPercent = 4;
        const int ColTotal = 5;
        const int ColPassed = 6;

        ParsedData parsedData;
        int maxColumnUsed;
        IConditionalFormattingRule[] failedFormattingRules;

        public XLSXTestResultSummarySheetWriter(ISheet sheet)
            : base(sheet)
        {
            MakeFailedFormattingRules();
        }

        public void CreateSheet(int yellowBand, int greenBand, ParsedData parsedData)
        {
            SetPercentageBands(yellowBand, greenBand);
            this.parsedData = parsedData.ThrowIfNull("parsedData");
            maxColumnUsed = ColTotal + parsedData.OutcomeNames.Count();

            int rowNum = CreateSummary(ByTestAssembly, 0, parsedData.SummaryByAssembly, parsedData.OutcomeNames);
            rowNum++;
            rowNum = CreateSummary(ByTestClass, rowNum, parsedData.SummaryByClass, parsedData.OutcomeNames);
            rowNum++;
            CreateSlowestTests(rowNum, parsedData.SlowestTests);

            sheet.SetColumnWidths(ColAssembly, 12000, 12000, 4000, 4000, 2500);
            for (int colNum = ColTotal; colNum <= maxColumnUsed; colNum++)
            {
                sheet.SetColumnWidth(colNum, 3000);
            }
        }

        int CreateSummary(string largeHeaderName, int rowNum, TestResultSummary summary, IEnumerable<string> outcomes)
        {
            int colNum = 0;
            int topOfSums = rowNum + 3;

            rowNum = CreateSummaryHeading(largeHeaderName, rowNum);

            IRow row = sheet.CreateRow(rowNum++);
            row.CreateHeadings(ColAssembly, "Assembly", "Class", "Time (secs)", "Time (hh:mm:ss)", "Pass Rate", "Total");
            colNum = ColPassed;
            foreach (var oc in outcomes)
            {
                row.SetCell(colNum++, oc).HeaderStyle().ApplyStyle();
            }

            foreach (var s in summary.SummaryLines.OrderBy(line => line.AssemblyFileName).ThenBy(line => line.ClassName))
            {
                row = sheet.CreateRow(rowNum++);
                row.SetCell(ColAssembly, s.AssemblyFileName);
                row.SetCell(ColClass, s.ClassName);
                row.SetCell(ColTime, s.TotalDurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(ColTimeHuman, s.TotalDurationInSecondsHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(ColPercent, s.PercentPassed).FormatPercentage().ApplyStyle();
                row.SetCell(ColTotal, s.TotalTests);

                colNum = ColPassed;
                foreach (var oc in s.Outcomes)
                {
                    var cell = row.SetCell(colNum++, oc.NumTests);
                    if (oc.Outcome == KnownOutcomes.Failed && oc.NumTests > 0)
                        cell.SolidFillColor(HSSFColor.Red.Index).ApplyStyle();
                    else if (oc.Outcome != KnownOutcomes.Passed && oc.NumTests > 0)
                        cell.SolidFillColor(HSSFColor.Yellow.Index).ApplyStyle();
                }
            }

            row = sheet.CreateRow(rowNum++);
            row.SetCell(ColTime, summary.TotalDurationInSeconds).SummaryStyle().Format("0.00").ApplyStyle();
            row.SetCell(ColTimeHuman, summary.TotalDurationInSecondsHuman).SummaryStyle().Alignment(HorizontalAlignment.Right).ApplyStyle();
            row.SetCell(ColPercent, summary.PercentPassed).SummaryStyle().FormatPercentage().ApplyStyle();
            row.SetCell(ColTotal, summary.TotalTests).SummaryStyle().ApplyStyle();

            int cn = ColPassed;
            foreach (var oc in outcomes)
            {
                row.SetCell(cn++, summary.TotalByOutcome(oc)).SummaryStyle().ApplyStyle();
            }

            ApplyPercentageFormatting(ColPercent, topOfSums, rowNum - 1);
            ApplyFailedFormattingRules(topOfSums, rowNum - 1);

            return rowNum;
        }

        void CreateSlowestTests(int rowNum, IEnumerable<SlowestTest> slowestTests)
        {
            rowNum = CreateSummaryHeading("20 Slowest Tests", rowNum);
            IRow row = sheet.CreateRow(rowNum++);
            row.CreateHeadings(ColAssembly, "Assembly", "Class", "Time (secs)", "Time (hh:mm:ss)", "Test Name");
            var range = new CellRangeAddress(row.RowNum, row.RowNum, ColPercent, maxColumnUsed);
            sheet.AddMergedRegion(range);

            foreach (var t in slowestTests)
            {
                row = sheet.CreateRow(rowNum++);
                row.SetCell(ColAssembly, t.AssemblyFileName);
                row.SetCell(ColClass, t.ClassName);
                row.SetCell(ColTime, t.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(ColTimeHuman, t.DurationHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(ColPercent, t.TestName);
                range = new CellRangeAddress(row.RowNum, row.RowNum, ColPercent, maxColumnUsed);
                sheet.AddMergedRegion(range);
            }
        }

        int CreateSummaryHeading(string largeHeaderName, int rowNum)
        {
            IRow row = sheet.CreateRow(rowNum++);
            row.SetCell(0, largeHeaderName).LargeHeaderStyle().ApplyStyle();
            row.HeightInPoints = 30;
            var range = new CellRangeAddress(row.RowNum, row.RowNum, 0, maxColumnUsed);
            sheet.AddMergedRegion(range);
            return rowNum;
        }

        void ApplyFailedFormattingRules(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("I{0}:I{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, failedFormattingRules);
        }

        void MakeFailedFormattingRules()
        {
            IConditionalFormattingRule rule1 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.NotEqual, "0");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = IndexedColors.Red.Index;
            fill1.FillPattern = (short)FillPattern.SolidForeground;

            failedFormattingRules = new IConditionalFormattingRule[] { rule1 };
        }
    }
}
