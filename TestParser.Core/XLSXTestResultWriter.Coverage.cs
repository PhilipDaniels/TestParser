using System;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TestParser.Core
{
    public partial class XLSXTestResultWriter
    {
        const int ColProjectFileName = 0;
        const int ColSourceFileName = 1;
        const int ColCoverage = 2;
        const int ColCompiledLines = 3;
        const int ColCoveredLines = 4;
        const int ColUncoveredLines = 5;
        const int ColProjectPathName = 6;
        const int ColSourcePathName = 7;

        ISheet coverageSheet;
        IConditionalFormattingRule[] coveragePercentageFormattingRules;

        void CreateCoverageSheet(ISheet sheet)
        {
            this.coverageSheet = sheet;
            coveragePercentageFormattingRules = MakePercentageConditionalFormattingRules(coverageSheet);

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
            foreach (var r in testResults.CoverageData.SortedByName)
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
            coverageSheet.SetColumnWidth(ColCoverage, 4000);
            coverageSheet.SetColumnWidth(ColCompiledLines, 4000);
            coverageSheet.SetColumnWidth(ColCoveredLines, 4000);
            coverageSheet.SetColumnWidth(ColUncoveredLines, 4000);
            coverageSheet.SetColumnWidth(ColProjectPathName, 10000);
            coverageSheet.SetColumnWidth(ColSourcePathName, 10000);

            ApplyPercentageFormatting(coverageSheet, 1, i);
        }

        void ApplyPercentageFormatting(ISheet sheet, int rowFromInclusive, int rowtoInclusive)
        {
            string range = String.Format("C{0}:C{1}", rowFromInclusive, rowtoInclusive);
            var region = new CellRangeAddress[] { CellRangeAddress.ValueOf(range) };
            sheet.SheetConditionalFormatting.AddConditionalFormatting(region, coveragePercentageFormattingRules);
        }
    }
}
