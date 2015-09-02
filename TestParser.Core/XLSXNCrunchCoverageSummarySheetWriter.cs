using BassUtils;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;

namespace TestParser.Core
{
    public class XLSXNCrunchCoverageSummarySheetWriter : XLSXSheetWriterBase
    {
        const int ColProjectFileName = 0;
        const int ColCoverage = 1;
        const int ColCompiledLines = 2;
        const int ColCoveredLines = 3;
        const int ColUncoveredLines = 4;
        const int ColProjectPathName = 5;

        public XLSXNCrunchCoverageSummarySheetWriter(ISheet sheet)
            : base(sheet)
        {
        }

        public void CreateSheet(int yellowBand, int greenBand, NCrunchCoverageDataCollection coverage)
        {
            SetPercentageBands(yellowBand, greenBand);
            coverage.ThrowIfNull("coverageData");

            IRow row = sheet.CreateRow(0);
            row.CreateHeadings(ColProjectFileName, "ProjectFileName", "Coverage", "CompiledLines",
                "CoveredLines", "UncoveredLines", "ProjectPathName");

            sheet.SetColumnWidths(ColProjectFileName, 10000, 4000, 4000, 4000, 4000, 20000);

            int i = 1;
            foreach (var r in coverage.SummariseByProject)
            {
                row = sheet.CreateRow(i);
                row.SetCell(ColProjectFileName, r.ProjectFileName);
                row.SetCell(ColCoverage, r.Coverage).FormatPercentage().ApplyStyle();
                row.SetCell(ColCompiledLines, r.CompiledLines);
                row.SetCell(ColCoveredLines, r.CoveredLines);
                row.SetCell(ColUncoveredLines, r.UncoveredLines);
                row.SetCell(ColProjectPathName, r.ProjectPathName);

                i++;
            }

            sheet.FreezeTopRow();
            ApplyPercentageFormatting(ColCoverage, 1, i);
        }
    }
}
