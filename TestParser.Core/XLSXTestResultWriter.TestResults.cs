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
    public partial class XLSXTestResultWriter
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

        ISheet resultsSheet;

        void CreateResultsSheet(ISheet sheet)
        {
            this.resultsSheet = sheet;

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
            resultsSheet.SheetConditionalFormatting.AddConditionalFormatting(region, ResultOutcomeFormattingRules);
        }

        IConditionalFormattingRule[] ResultOutcomeFormattingRules
        {
            get
            {
                if (resultOutcomeFormattingRules == null)
                {
                    IConditionalFormattingRule rule1 = resultsSheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Passed + "\"");
                    IPatternFormatting fill1 = rule1.CreatePatternFormatting();
                    fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
                    fill1.FillPattern = (short)FillPattern.SolidForeground;

                    IConditionalFormattingRule rule2 = resultsSheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Failed + "\"");
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
