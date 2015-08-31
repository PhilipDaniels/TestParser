using System;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public partial class XLSXTestResultWriter
    {
        const int ColCovSummaryProjectFileName = 0;
        const int ColCovSummaryCoverage = 1;
        const int ColCovSummaryCompiledLines = 2;
        const int ColCovSummaryCoveredLines = 3;
        const int ColCovSummaryUncoveredLines = 4;
        const int ColCovSummaryProjectPathName = 5;

        ISheet coverageSummarySheet;
        IConditionalFormattingRule[] coverageSummaryPercentageFormattingRules;

        private void CreateCoverageSummarySheet(ISheet sheet)
        {
            this.coverageSummarySheet = sheet;
            coverageSummaryPercentageFormattingRules = MakePercentageConditionalFormattingRules(coverageSummarySheet);

            IRow row = coverageSummarySheet.CreateRow(0);
            row.SetCell(ColCovSummaryProjectFileName, "ProjectFileName").HeaderStyle().ApplyStyle();
            row.SetCell(ColCovSummaryCoverage, "Coverage").HeaderStyle().ApplyStyle();
            row.SetCell(ColCovSummaryCompiledLines, "CompiledLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColCovSummaryCoveredLines, "CoveredLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColCovSummaryUncoveredLines, "UncoveredLines").HeaderStyle().ApplyStyle();
            row.SetCell(ColCovSummaryProjectPathName, "ProjectPathName").HeaderStyle().ApplyStyle();

            int i = 1;
            foreach (var r in testResults.CoverageData.SummariseByProject)
            {
                row = coverageSummarySheet.CreateRow(i);
                row.SetCell(ColProjectFileName, r.ProjectFileName);
                row.SetCell(ColCovSummaryCoverage, r.Coverage).FormatPercentage().ApplyStyle();
                row.SetCell(ColCovSummaryCompiledLines, r.CompiledLines);
                row.SetCell(ColCovSummaryCoveredLines, r.CoveredLines);
                row.SetCell(ColCovSummaryUncoveredLines, r.UncoveredLines);
                row.SetCell(ColCovSummaryProjectPathName, r.ProjectPathName);

                i++;
            }

            // Freeze the header row.
            coverageSummarySheet.CreateFreezePane(0, 1, 0, 1);

            coverageSummarySheet.SetColumnWidth(ColCovSummaryProjectFileName, 10000);
            coverageSummarySheet.SetColumnWidth(ColCovSummaryCoverage, 4000);
            coverageSummarySheet.SetColumnWidth(ColCovSummaryCompiledLines, 4000);
            coverageSummarySheet.SetColumnWidth(ColCovSummaryCoveredLines, 4000);
            coverageSummarySheet.SetColumnWidth(ColCovSummaryUncoveredLines, 4000);
            coverageSummarySheet.SetColumnWidth(ColCovSummaryProjectPathName, 20000);

            ApplyCovSummaryPercentageFormatting(coverageSummarySheet, 1, i);
        }

        void ApplyCovSummaryPercentageFormatting(ISheet sheet, int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("B{0}:B{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, coverageSummaryPercentageFormattingRules);
        }
    }
}
