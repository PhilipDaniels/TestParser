using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace TestParser.Core
{
    public class XLSXTestResultWriter : ITestResultWriter
    {
        IWorkbook workbook;
        ISheet summarySheet;
        ISheet resultsSheet;
        ISheet coverageSheet;
        ISheetConditionalFormatting summarySheetConditionalFormatting;
        ISheetConditionalFormatting resultsSheetConditionalFormatting;
        TestResults testResults;
        readonly string yellowBandString;
        readonly string greenBandString;

        const int SumColAssembly = 0;
        const int SumColClass = 1;
        const int SumColTime = 2;
        const int SumColTimeHuman = 3;
        const int SumColPercent = 4;
        const int SumColTotal = 5;
        const int SumColPassed = 6;

        public XLSXTestResultWriter(int yellowBand, int greenBand)
        {
            yellowBandString = (((decimal)yellowBand) / 100m).ToString();
            greenBandString = (((decimal)greenBand) / 100m).ToString();
        }

        public void WriteResults(Stream s, TestResults testResults)
        {
            workbook = new XSSFWorkbook();
            summarySheet = workbook.CreateSheet("Summary");
            resultsSheet = workbook.CreateSheet("TestResults");
            coverageSheet = workbook.CreateSheet("CoverageData");

            this.testResults = testResults;
            summarySheetConditionalFormatting = summarySheet.SheetConditionalFormatting;
            resultsSheetConditionalFormatting = resultsSheet.SheetConditionalFormatting;

            CreateCoverageSheet();
            CreateResultsSheet();
            CreateSummarySheet();

            workbook.Write(s);

            // Handy check to see if caching is working.
            int numStylesCreated = FluentCell.NumCachedStyles;
        }

        void CreateCoverageSheet()
        {
            const int ColProjectFileName = 0;
            const int ColSourceFileName = 1;
            const int ColCoverage = 2;
            const int ColCompiledLines = 3;
            const int ColCoveredLines = 4;
            const int ColUncoveredLines = 5;
            const int ColProjectPathName = 6;
            const int ColSourcePathName = 7;

            IRow row = coverageSheet.CreateRow(0);
            row.SetCell(ColProjectFileName, "ProjectFileName").HeaderStyle().ApplyStyle();
            row.SetCell(ColSourceFileName, "SourceFileName").HeaderStyle().ApplyStyle();
            row.SetCell(ColCoverage, "Coverage").HeaderStyle().ApplyStyle();
            row.SetCell(ColCompiledLines, "CompiledLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColCoveredLines, "CoveredLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColUncoveredLines, "UncoveredLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColProjectPathName, "ProjectPathName").HeaderStyle().ApplyStyle();
            row.SetCell(ColSourcePathName, "SourcePathName").HeaderStyle().ApplyStyle();

            int i = 1;
            foreach (var r in testResults.CoverageData)
            {
                row = coverageSheet.CreateRow(i);
                row.SetCell(ColProjectFileName, r.ProjectFileName);
                row.SetCell(ColSourceFileName, r.SourceFileName);
                row.SetCell(ColCoverage, r.Coverage).FormatPercentage().ApplyStyle();
                row.SetCell(ColCompiledLines, r.CompiledLines);
                row.SetCell(ColCoveredLines, r.CoveredLines);
                row.SetCell(ColUncoveredLines, r.UncoveredLines);
                row.SetCell(ColProjectPathName, r.ProjectPathName);
                row.SetCell(ColSourcePathName, r.SourceFilePathName);

                i++;
            }

            // Freeze the header row.
            coverageSheet.CreateFreezePane(0, 1, 0, 1);

            coverageSheet.SetColumnWidth(ColProjectFileName, 10000);
            coverageSheet.SetColumnWidth(ColSourceFileName, 10000);
            coverageSheet.SetColumnWidth(ColCoverage, 3000);
            coverageSheet.SetColumnWidth(ColCompiledLines, 3000);
            coverageSheet.SetColumnWidth(ColCoveredLines, 3000);
            coverageSheet.SetColumnWidth(ColUncoveredLines, 3000);
            coverageSheet.SetColumnWidth(ColProjectPathName, 10000);
            coverageSheet.SetColumnWidth(ColSourcePathName, 10000);
        }

        void CreateResultsSheet()
        {
            const int ColAssemblyFileName = 0;
            const int ColClassName = 1;
            const int ColTestName = 2;
            const int ColOutcome = 3;
            const int ColDurationInSeconds = 4;
            const int ColDurationInSecondsHuman = 5;
            const int ColErrorMessage = 6;
            const int ColStackTrace = 7;
            const int ColStartTime = 8;
            const int ColEndTime = 9;
            const int ColResultsPathName = 10;
            const int ColResultsFileName = 11;
            const int ColAssemblyPathName = 12;
            const int ColFullClassName = 13;
            const int ColComputerName = 14;
            const int ColTestResultFileType = 15;

            IRow row = resultsSheet.CreateRow(0);
            row.SetCell(ColResultsPathName, "ResultsPathName").HeaderStyle().ApplyStyle();
            row.SetCell(ColResultsFileName, "ResultsFileName").HeaderStyle().ApplyStyle();
            row.SetCell(ColAssemblyPathName, "AssemblyPathName").HeaderStyle().ApplyStyle();
            row.SetCell(ColAssemblyFileName, "AssemblyFileName").HeaderStyle().ApplyStyle();
            row.SetCell(ColFullClassName, "FullClassName").HeaderStyle().ApplyStyle();
            row.SetCell(ColClassName, "ClassName").HeaderStyle().ApplyStyle();
            row.SetCell(ColTestName, "TestName").HeaderStyle().ApplyStyle();
            row.SetCell(ColComputerName, "ComputerName").HeaderStyle().ApplyStyle();
            row.SetCell(ColStartTime, "StartTime").HeaderStyle().ApplyStyle();
            row.SetCell(ColEndTime, "EndTime").HeaderStyle().ApplyStyle();
            row.SetCell(ColDurationInSeconds, "DurationInSeconds").HeaderStyle().ApplyStyle();
            row.SetCell(ColDurationInSecondsHuman, "DurationInSecondsHuman").HeaderStyle().Alignment(HorizontalAlignment.Right).ApplyStyle();
            row.SetCell(ColOutcome, "Outcome").HeaderStyle().ApplyStyle();
            row.SetCell(ColErrorMessage, "ErrorMessage").HeaderStyle().ApplyStyle();
            row.SetCell(ColStackTrace, "StackTrace").HeaderStyle().ApplyStyle();
            row.SetCell(ColTestResultFileType, "TestResultFileType").HeaderStyle().ApplyStyle();

            int i = 1;
            foreach (var r in testResults.ResultLines.SortedByFailedOtherPassed)
            {
                row = resultsSheet.CreateRow(i);
                row.SetCell(ColResultsPathName, r.ResultsPathName);
                row.SetCell(ColResultsFileName, r.ResultsFileName);
                row.SetCell(ColAssemblyPathName, r.AssemblyPathName);
                row.SetCell(ColAssemblyFileName, r.AssemblyFileName);
                row.SetCell(ColFullClassName, r.FullClassName);
                row.SetCell(ColClassName, r.ClassName);
                row.SetCell(ColTestName, r.TestName);
                row.SetCell(ColComputerName, r.ComputerName);
                if (r.StartTime != null)
                    row.SetCell(ColStartTime, r.StartTime.Value).FormatLongDate().ApplyStyle();
                if (r.EndTime != null)
                    row.SetCell(ColEndTime, r.EndTime.Value).FormatLongDate().ApplyStyle();
                row.SetCell(ColDurationInSeconds, r.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(ColDurationInSecondsHuman, r.DurationHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(ColOutcome, r.Outcome);
                row.SetCell(ColErrorMessage, r.ErrorMessage);
                row.SetCell(ColStackTrace, r.StackTrace);
                row.SetCell(ColTestResultFileType, r.TestResultFileType.ToString());

                i++;
            }

            // Freeze the header row.
            resultsSheet.CreateFreezePane(0, 1, 0, 1);

            resultsSheet.SetColumnWidth(ColAssemblyFileName, 10000);
            resultsSheet.SetColumnWidth(ColClassName, 10000);
            resultsSheet.SetColumnWidth(ColTestName, 10000);
            resultsSheet.SetColumnWidth(ColOutcome, 3000);
            resultsSheet.SetColumnWidth(ColDurationInSeconds, 3000);
            resultsSheet.SetColumnWidth(ColErrorMessage, 3000);
            resultsSheet.SetColumnWidth(ColStackTrace, 3000);
            resultsSheet.SetColumnWidth(ColStartTime, 5500);
            resultsSheet.SetColumnWidth(ColEndTime, 5500);

            string range = String.Format("D2:D{0}", i);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            resultsSheetConditionalFormatting.AddConditionalFormatting(region, ResultOutcomeFormattingRules);
        }

        void CreateSummarySheet()
        {
            int rowNum = CreateSummary("Summary By Assembly", 0, testResults.SummaryByAssembly, testResults.OutcomeNames);
            rowNum++;
            rowNum = CreateSummary("Summary By Class", rowNum, testResults.SummaryByClass, testResults.OutcomeNames);
            rowNum++;
            CreateSlowestTests(rowNum, testResults.SlowestTests);

            // Set the widths.
            summarySheet.SetColumnWidth(SumColAssembly, 12000);
            summarySheet.SetColumnWidth(SumColClass, 12000);
            summarySheet.SetColumnWidth(SumColTime, 4000);
            summarySheet.SetColumnWidth(SumColTimeHuman, 4000);
            summarySheet.SetColumnWidth(SumColPercent, 2500);
            for (int colNum = SumColTotal; colNum <= MaxSummaryColumnUsed; colNum++)
                summarySheet.SetColumnWidth(colNum, 3000);
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
            rowNum = CreateSummaryHeading("10 Slowest Tests", rowNum);
            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(SumColAssembly, "Assembly").HeaderStyle().ApplyStyle();
            row.SetCell(SumColClass, "Class").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTime, "Time (secs)").HeaderStyle().ApplyStyle();
            row.SetCell(SumColTimeHuman, "Time (hh:mm:ss)").HeaderStyle().ApplyStyle();

            row.SetCell(SumColPercent, "Test Name").HeaderStyle().ApplyStyle();
            var range = new CellRangeAddress(row.RowNum, row.RowNum, SumColPercent, MaxSummaryColumnUsed);
            summarySheet.AddMergedRegion(range);

            foreach (var t in slowestTests)
            {
                row = summarySheet.CreateRow(rowNum++);
                row.SetCell(SumColAssembly, t.AssemblyFileName);
                row.SetCell(SumColClass, t.ClassName);
                row.SetCell(SumColTime, t.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(SumColTimeHuman, t.DurationHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(SumColPercent, t.TestName);
                range = new CellRangeAddress(row.RowNum, row.RowNum, SumColPercent, MaxSummaryColumnUsed);
                summarySheet.AddMergedRegion(range);
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
            row.SetCell(SumColPercent, "Percent").HeaderStyle().ApplyStyle();
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
            row.SetCell(SumColPercent, summary.PercentPassed).SummaryStyle().FormatPercentage().ApplyStyle();
            row.SetCell(SumColTotal, summary.TotalTests).SummaryStyle().ApplyStyle();

            int cn = SumColPassed;
            foreach (var oc in outcomes)
            {
                row.SetCell(cn++, summary.TotalByOutcome(oc)).SummaryStyle().ApplyStyle();
            }

            ConditionalFormatPercentage(topOfSums, rowNum - 1);
            ConditionalFormatFails(topOfSums, rowNum - 1);

            return rowNum;
        }

        private int CreateSummaryHeading(string largeHeaderName, int rowNum)
        {
            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(0, largeHeaderName).LargeHeaderStyle().ApplyStyle();
            row.HeightInPoints = 30;
            var range = new CellRangeAddress(row.RowNum, row.RowNum, 0, MaxSummaryColumnUsed);
            summarySheet.AddMergedRegion(range);
            return rowNum;
        }

        void ConditionalFormatPercentage(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("E{0}:E{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheetConditionalFormatting.AddConditionalFormatting(region, SummaryPercentageFormattingRules);
        }

        void ConditionalFormatFails(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("H{0}:H{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheetConditionalFormatting.AddConditionalFormatting(region, SummaryFailedFormattingRules);
        }

        IConditionalFormattingRule[] SummaryPercentageFormattingRules
        {
            get
            {
                if (summaryPercentageFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, greenBandString);
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule2 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, yellowBandString);
                    IPatternFormatting fill2 = rule2.CreatePatternFormatting();
                    fill2.FillBackgroundColor = IndexedColors.Yellow.Index;
                    fill2.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule3 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, yellowBandString);
                    IPatternFormatting fill3 = rule3.CreatePatternFormatting();
                    fill3.FillBackgroundColor = IndexedColors.Red.Index;
                    fill3.FillPattern = (short)FillPattern.SolidForeground;

                    summaryPercentageFormattingRules = new IConditionalFormattingRule[] { rule1, rule2, rule3 };
                }

                return summaryPercentageFormattingRules;
            }
        }
        IConditionalFormattingRule[] summaryPercentageFormattingRules;

        IConditionalFormattingRule[] SummaryFailedFormattingRules
        {
            get
            {
                if (summaryFailedFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.NotEqual, "0");
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.Red.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    summaryFailedFormattingRules = new IConditionalFormattingRule[] { rule1 };
                }

                return summaryFailedFormattingRules;
            }
        }
        IConditionalFormattingRule[] summaryFailedFormattingRules;

        IConditionalFormattingRule[] ResultOutcomeFormattingRules
        {
            get
            {
                if (resultOutcomeFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Passed + "\"");
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule2 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Failed + "\"");
                    IPatternFormatting fill2 = rule2.CreatePatternFormatting();
                    fill2.FillBackgroundColor = IndexedColors.Red.Index;
                    fill2.FillPattern = (short)FillPattern.SolidForeground;

                    resultOutcomeFormattingRules = new IConditionalFormattingRule[] { rule1, rule2 };
                }

                return resultOutcomeFormattingRules;
            }
        }
        IConditionalFormattingRule[] resultOutcomeFormattingRules;
    }
}
