﻿using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public partial class XLSXTestResultWriter
    {
        const string ByTestAssembly = "Summary By Test Assembly";
        const string ByTestClass = "Summary By Test Class";

        const int SumColAssembly = 0;
        const int SumColClass = 1;
        const int SumColTime = 2;
        const int SumColTimeHuman = 3;
        const int SumColCoverage = 4;
        const int SumColPercent = 5;
        const int SumColTotal = 6;
        const int SumColPassed = 7;

        ISheet summarySheet;
        IConditionalFormattingRule[] summaryPercentageFormattingRules;

        void CreateSummarySheet(ISheet sheet)
        {
            this.summarySheet = sheet;
            summaryPercentageFormattingRules = MakePercentageConditionalFormattingRules(summarySheet);

            int rowNum = CreateSummary(ByTestAssembly, 0, testResults.SummaryByAssembly, testResults.OutcomeNames);
            rowNum++;
            rowNum = CreateSummary(ByTestClass, rowNum, testResults.SummaryByClass, testResults.OutcomeNames);
            rowNum++;
            CreateSlowestTests(rowNum, testResults.SlowestTests);

            // Set the widths.
            summarySheet.SetColumnWidth(SumColAssembly, 12000);
            summarySheet.SetColumnWidth(SumColClass, 12000);
            summarySheet.SetColumnWidth(SumColTime, 4000);
            summarySheet.SetColumnWidth(SumColTimeHuman, 4000);
            summarySheet.SetColumnWidth(SumColCoverage, 2500);
            summarySheet.SetColumnWidth(SumColPercent, 2500);
            for (int colNum = SumColTotal; colNum <= MaxSummaryColumnUsed; colNum++)
            {
                summarySheet.SetColumnWidth(colNum, 3000);
            }
        }

        int CreateSummary(string largeHeaderName, int rowNum, TestResultSummary summary, IEnumerable<string> outcomes)
        {
            int colNum = 0;
            int topOfSums = rowNum + 3;

            rowNum = CreateSummaryHeading(largeHeaderName, rowNum);

            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(SumColAssembly, "Assembly").HeaderStyle().ApplyStyle();
            row.SetCell(SumColClass, "Class").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTime, "Time (secs)").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTimeHuman, "Time (hh:mm:ss)").HeaderStyle().ApplyStyle();
            row.SetCell(SumColCoverage, "Coverage").HeaderStyle().ApplyStyle();
            row.SetCell(SumColPercent, "Pass Rate").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTotal, "Total").HeaderStyle().ApplyStyle();
            colNum = SumColPassed;
            foreach (var oc in outcomes)
            {
                row.SetCell(colNum++, oc).HeaderStyle().ApplyStyle();
            }

            foreach (var s in summary.SummaryLines.OrderBy(line => line.AssemblyFileName).ThenBy(line => line.ClassName))
            {
                row = summarySheet.CreateRow(rowNum++);
                row.SetCell(SumColAssembly, s.AssemblyFileName);
                row.SetCell(SumColClass, s.ClassName);
                row.SetCell(SumColTime, s.TotalDurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(SumColTimeHuman, s.TotalDurationInSecondsHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                if (largeHeaderName == ByTestAssembly)
                    row.SetCell(SumColCoverage, s.Coverage).FormatPercentage().ApplyStyle();
                row.SetCell(SumColPercent, s.PercentPassed).FormatPercentage().ApplyStyle();
                row.SetCell(SumColTotal, s.TotalTests);

                colNum = SumColPassed;
                foreach (var oc in s.Outcomes)
                {
                    var cell = row.SetCell(colNum++, oc.NumTests);
                    if (oc.Outcome == KnownOutcomes.Failed && oc.NumTests > 0)
                        cell.SolidFillColor(HSSFColor.Red.Index).ApplyStyle();
                    else if (oc.Outcome != KnownOutcomes.Passed && oc.NumTests > 0)
                        cell.SolidFillColor(HSSFColor.Yellow.Index).ApplyStyle();
                }
            }

            row = summarySheet.CreateRow(rowNum++);
            row.SetCell(SumColTime, summary.TotalDurationInSeconds).SummaryStyle().Format("0.00").ApplyStyle();
            row.SetCell(SumColTimeHuman, summary.TotalDurationInSecondsHuman).SummaryStyle().Alignment(HorizontalAlignment.Right).ApplyStyle();
            if (largeHeaderName == ByTestAssembly)
                row.SetCell(SumColCoverage, summary.Coverage).SummaryStyle().FormatPercentage().ApplyStyle();
            else
                row.SetCell(SumColCoverage, "").SummaryStyle().ApplyStyle();
            row.SetCell(SumColPercent, summary.PercentPassed).SummaryStyle().FormatPercentage().ApplyStyle();
            row.SetCell(SumColTotal, summary.TotalTests).SummaryStyle().ApplyStyle();

            int cn = SumColPassed;
            foreach (var oc in outcomes)
            {
                row.SetCell(cn++, summary.TotalByOutcome(oc)).SummaryStyle().ApplyStyle();
            }

            if (largeHeaderName == ByTestAssembly)
                ApplySummaryCoverageFormatting(topOfSums, rowNum - 1);
            ApplySummaryPercentageFormatting(topOfSums, rowNum - 1);
            ApplySummaryFailsFormatting(topOfSums, rowNum - 1);

            return rowNum;
        }

        int CreateSummaryHeading(string largeHeaderName, int rowNum)
        {
            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(0, largeHeaderName).LargeHeaderStyle().ApplyStyle();
            row.HeightInPoints = 30;
            var range = new CellRangeAddress(row.RowNum, row.RowNum, 0, MaxSummaryColumnUsed);
            summarySheet.AddMergedRegion(range);
            return rowNum;
        }

        /// <summary>
        /// Gets the maximum summary column used.
        /// </summary>
        /// <value>
        /// The maximum summary column used.
        /// </value>
        int MaxSummaryColumnUsed
        {
            get
            {
                return SumColTotal + testResults.OutcomeNames.Count();
            }
        }

        void CreateSlowestTests(int rowNum, IEnumerable<SlowestTest> slowestTests)
        {
            rowNum = CreateSummaryHeading("20 Slowest Tests", rowNum);
            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(SumColAssembly, "Assembly").HeaderStyle().ApplyStyle();
            row.SetCell(SumColClass, "Class").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTime, "Time (secs)").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTimeHuman, "Time (hh:mm:ss)").HeaderStyle().ApplyStyle();
            row.SetCell(SumColCoverage, "Test Name").HeaderStyle().ApplyStyle();
            var range = new CellRangeAddress(row.RowNum, row.RowNum, SumColCoverage, MaxSummaryColumnUsed);
            summarySheet.AddMergedRegion(range);

            foreach (var t in slowestTests)
            {
                row = summarySheet.CreateRow(rowNum++);
                row.SetCell(SumColAssembly, t.AssemblyFileName);
                row.SetCell(SumColClass, t.ClassName);
                row.SetCell(SumColTime, t.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(SumColTimeHuman, t.DurationHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(SumColCoverage, t.TestName);
                range = new CellRangeAddress(row.RowNum, row.RowNum, SumColCoverage, MaxSummaryColumnUsed);
                summarySheet.AddMergedRegion(range);
            }
        }

        void ApplySummaryCoverageFormatting(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("E{0}:E{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheet.SheetConditionalFormatting.AddConditionalFormatting(region, summaryPercentageFormattingRules);
        }

        void ApplySummaryPercentageFormatting(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("F{0}:F{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheet.SheetConditionalFormatting.AddConditionalFormatting(region, summaryPercentageFormattingRules);
        }

        void ApplySummaryFailsFormatting(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("I{0}:I{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheet.SheetConditionalFormatting.AddConditionalFormatting(region, SummaryFailedFormattingRules);
        }

        IConditionalFormattingRule[] SummaryFailedFormattingRules
        {
            get
            {
                if (summaryFailedFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = summarySheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.NotEqual, "0");
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.Red.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    summaryFailedFormattingRules = new IConditionalFormattingRule[] { rule1 };
                }

                return summaryFailedFormattingRules;
            }
        }
        IConditionalFormattingRule[] summaryFailedFormattingRules;
    }
}
