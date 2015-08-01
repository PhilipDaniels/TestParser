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
        ISheetConditionalFormatting summarySheetConditionalFormatting;
        ISheetConditionalFormatting resultsSheetConditionalFormatting;

        public void WriteResults(Stream s, IEnumerable<TestResult> testResults)
        {
            workbook = new XSSFWorkbook();
            summarySheet = workbook.CreateSheet("Summary");
            resultsSheet = workbook.CreateSheet("TestResults");
            summarySheetConditionalFormatting = summarySheet.SheetConditionalFormatting;
            resultsSheetConditionalFormatting = resultsSheet.SheetConditionalFormatting;

            CreateResultsSheet(testResults);
            CreateSummarySheet(testResults);

            workbook.Write(s);

            // Handy check to see caching is working.
            int numStylesCreated = FluentCell.NumCachedStyles;
        }

        void CreateResultsSheet(IEnumerable<TestResult> testResults)
        {
            const int ColAssemblyFileName = 0;
            const int ColClassName = 1;
            const int ColTestName = 2;
            const int ColOutcome = 3;
            const int ColDurationInSeconds = 4;
            const int ColErrorMessage = 5;
            const int ColStackTrace = 6;
            const int ColStartTime = 7;
            const int ColEndTime = 8;
            const int ColResultsPathName = 9;
            const int ColResultsFileName = 10;
            const int ColAssemblyPathName = 11;
            const int ColFullClassName = 12;
            const int ColComputerName = 13;
            const int ColTestResultFileType = 14;

            IRow row = resultsSheet.CreateRow(0);
            row.SetCell(ColResultsPathName, "ResultsPathName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColResultsFileName, "ResultsFileName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColAssemblyPathName, "AssemblyPathName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColAssemblyFileName, "AssemblyFileName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColFullClassName, "FullClassName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColClassName, "ClassName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColTestName, "TestName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColComputerName, "ComputerName").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColStartTime, "StartTime").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColEndTime, "EndTime").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColDurationInSeconds, "DurationInSeconds").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColOutcome, "Outcome").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColErrorMessage, "ErrorMessage").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColStackTrace, "StackTrace").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColTestResultFileType, "TestResultFileType").HeaderStyle().ApplyStyle();

            int i = 1;
            foreach (var r in testResults)
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

        void CreateSummarySheet(IEnumerable<TestResult> testResults)
        {
            var sba = TestResultSummary.SummariseByAssembly(testResults);
            int rowNum = CreateSummary("Summary By Assembly", 0, sba);

            var sbc = TestResultSummary.SummariseByClass(testResults);
            rowNum++;
            rowNum = CreateSummary("Summary By Class", rowNum, sbc);

            // Set the widths.
            summarySheet.SetColumnWidth(0, 12000);
            summarySheet.SetColumnWidth(1, 12000);
            summarySheet.SetColumnWidth(2, 3000);
            summarySheet.SetColumnWidth(3, 2500);
            for (int colNum = 4; colNum <= 4 + sba.ElementAt(0).Outcomes.Count(); colNum++)
                summarySheet.SetColumnWidth(colNum, 3000);
        }

        int CreateSummary(string largeHeaderName, int rowNum, IEnumerable<TestResultSummary> summary)
        {
            int colNum = 0;
            int topOfSums = rowNum + 3;

            IRow row = summarySheet.CreateRow(rowNum++);
            row.SetCell(0, largeHeaderName).LargeHeaderStyle().ApplyStyle();
            row.HeightInPoints = 30;
            var range = new CellRangeAddress(row.RowNum, row.RowNum, 0, 6);
            summarySheet.AddMergedRegion(range);

            row = summarySheet.CreateRow(rowNum++);
            row.SetCell(0, "Assembly").HeaderStyle().ApplyStyle();
            row.SetCell(1, "Class").HeaderStyle().ApplyStyle();
            row.SetCell(2, "Time (secs)").HeaderStyle().ApplyStyle();
            row.SetCell(3, "Percent").HeaderStyle().ApplyStyle();
            row.SetCell(4, "Total").HeaderStyle().ApplyStyle();
            colNum = 5;
            foreach (var oc in summary.ElementAt(0).Outcomes)
            {
                row.SetCell(colNum++, oc.Outcome).HeaderStyle().ApplyStyle();
            }

            foreach (var s in summary)
            {
                row = summarySheet.CreateRow(rowNum++);
                row.SetCell(0, s.AssemblyFileName);
                row.SetCell(1, s.ClassName);
                row.SetCell(2, s.TotalDurationInSeconds).Format("0.00").ApplyStyle();
                row.SetFormula(3, "F{0}/E{0}", rowNum).FormatPercentage().ApplyStyle();
                row.SetCell(4, s.TotalTests);

                colNum = 5;
                foreach (var oc in s.Outcomes)
                {
                    var cell = row.SetCell(colNum++, oc.NumTests);
                    if (oc.Outcome == KnownOutcomes.Failed && oc.NumTests > 0)
                        cell.SolidFillColor(HSSFColor.Red.Index).ApplyStyle();
                }
            }

            row = summarySheet.CreateRow(rowNum++);
            row.SetFormula(2, "SUM(C{0}:C{1})", topOfSums, rowNum - 1).SummaryStyle().Format("0.00").ApplyStyle();
            row.SetFormula(3, "F{0}/E{0}", rowNum).SummaryStyle().FormatPercentage().ApplyStyle();
            row.SetFormula(4, "SUM(E{0}:E{1})", topOfSums, rowNum - 1).SummaryStyle().ApplyStyle();
            int maxCol = 5 + summary.ElementAt(0).Outcomes.Count();
            for (int cn = 5; cn < maxCol; cn++)
            {
                string colRef = CellReference.ConvertNumToColString(cn);
                row.SetFormula(cn, "SUM({1}{0}:{1}{2})", topOfSums, colRef, rowNum - 1).SummaryStyle().ApplyStyle();
            }

            ConditionalFormatPercentage(topOfSums, rowNum - 1);
            ConditionalFormatFails(topOfSums, rowNum - 1);

            return rowNum;
        }

        void ConditionalFormatPercentage(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("D{0}:D{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheetConditionalFormatting.AddConditionalFormatting(region, SummaryPercentageFormattingRules);
        }

        void ConditionalFormatFails(int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("G{0}:G{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            summarySheetConditionalFormatting.AddConditionalFormatting(region, SummaryFailedFormattingRules);
        }

        IConditionalFormattingRule[] SummaryPercentageFormattingRules
        {
            get
            {
                if (summaryPercentageFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, "1");
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule2 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, "0.9");
                    IPatternFormatting fill2 = rule2.CreatePatternFormatting();
                    fill2.FillBackgroundColor = IndexedColors.Yellow.Index;
                    fill2.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule3 = summarySheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "0.9");
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
