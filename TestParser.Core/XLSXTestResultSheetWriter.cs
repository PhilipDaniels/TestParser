using System;
using BassUtils;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public class XLSXTestResultSheetWriter : XLSXSheetWriterBase
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

        IConditionalFormattingRule[] outcomeFormattingRules;

        public XLSXTestResultSheetWriter(ISheet sheet)
            : base(sheet)
        {
            MakeOutcomeFormattingRules();
        }

        public void CreateSheet(int yellowBand, int greenBand, ParsedData parsedData)
        {
            SetPercentageBands(yellowBand, greenBand);
            parsedData.ThrowIfNull("parsedData");

            IRow row = sheet.CreateRow(0);
            row.CreateHeadings(ColAssemblyFileName, "AssemblyFileName", "ClassName",
                "TestName", "Outcome", "DurationInSeconds", "DurationInSecondsHuman",
                "ErrorMessage", "StackTrace", "StartTime", "EndTime",
                "ResultsPathName", "ResultsFileName", "AssemblyPathName",
                "FullClassName", "ComputerName", "TestResultFileType");
            sheet.SetColumnWidths(ColAssemblyFileName, 10000, 10000, 10000, 3000, 3000, 3000, 3000, 5500, 5500);

            int i = 1;
            foreach (var r in parsedData.ResultLines.SortedByFailedOtherPassed)
            {
                row = sheet.CreateRow(i);
                row.SetCell(ColAssemblyFileName, r.AssemblyFileName);
                row.SetCell(ColClassName, r.ClassName);
                row.SetCell(ColTestName, r.TestName);
                row.SetCell(ColOutcome, r.Outcome);
                row.SetCell(ColDurationInSeconds, r.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(ColDurationInSecondsHuman, r.DurationHuman).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(ColErrorMessage, r.ErrorMessage);
                row.SetCell(ColStackTrace, r.StackTrace);
                if (r.StartTime != null)
                    row.SetCell(ColStartTime, r.StartTime.Value).FormatLongDate().ApplyStyle();
                if (r.EndTime != null)
                    row.SetCell(ColEndTime, r.EndTime.Value).FormatLongDate().ApplyStyle();
                row.SetCell(ColResultsPathName, r.ResultsPathName);
                row.SetCell(ColResultsFileName, r.ResultsFileName);
                row.SetCell(ColAssemblyPathName, r.AssemblyPathName);
                row.SetCell(ColFullClassName, r.FullClassName);
                row.SetCell(ColComputerName, r.ComputerName);
                row.SetCell(ColTestResultFileType, r.TestResultFileType.ToString());

                i++;
            }

            sheet.FreezeTopRow();
            ApplyOutcomeFormattingRules(i);
        }

        void ApplyOutcomeFormattingRules(int i)
        {
            string range = String.Format("D2:D{0}", i);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, outcomeFormattingRules);
        }

        void MakeOutcomeFormattingRules()
        {
            IConditionalFormattingRule rule1 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Passed + "\"");
            IPatternFormatting fill1 = rule1.CreatePatternFormatting();
            fill1.FillBackgroundColor = IndexedColors.BrightGreen.Index;
            fill1.FillPattern = (short)FillPattern.SolidForeground;

            IConditionalFormattingRule rule2 = sheet.SheetConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "\"" + KnownOutcomes.Failed + "\"");
            IPatternFormatting fill2 = rule2.CreatePatternFormatting();
            fill2.FillBackgroundColor = IndexedColors.Red.Index;
            fill2.FillPattern = (short)FillPattern.SolidForeground;

            outcomeFormattingRules = new IConditionalFormattingRule[] { rule1, rule2 };
        }
    }
}
