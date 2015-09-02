using BassUtils;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;

namespace TestParser.Core
{
    public class XLSXNCrunchCoverageSheetWriter : XLSXSheetWriterBase
    {
        const int ColProjectFileName = 0;
        const int ColSourceFileName = 1;
        const int ColCoverage = 2;
        const int ColCompiledLines = 3;
        const int ColCoveredLines = 4;
        const int ColUncoveredLines = 5;
        const int ColProjectPathName = 6;
        const int ColSourcePathName = 7;

        public XLSXNCrunchCoverageSheetWriter(ISheet sheet)
            : base(sheet)
        {
        }

        public void CreateSheet(int yellowBand, int greenBand, NCrunchCoverageDataCollection coverage)
        {
            SetPercentageBands(yellowBand, greenBand);
            coverage.ThrowIfNull("coverageData");

            IRow row = sheet.CreateRow(0);
            row.CreateHeadings(ColProjectFileName, "ProjectFileName", "SourceFileName",
                "Coverage", "CompiledLines", "CoveredLines",
                "UncoveredLines", "ProjectPathName", "SourcePathName"
                );

            sheet.SetColumnWidths(ColProjectFileName, 10000, 10000, 4000, 4000, 4000, 4000, 10000, 10000);

            int i = 1;
            foreach (var coverageRow in coverage.SortedByName)
            {
                row = sheet.CreateRow(i);
                row.SetCell(ColProjectFileName, coverageRow.ProjectFileName);
                row.SetCell(ColSourceFileName, coverageRow.SourceFileName);
                row.SetCell(ColCoverage, coverageRow.Coverage).FormatPercentage().ApplyStyle();
                row.SetCell(ColCompiledLines, coverageRow.CompiledLines);
                row.SetCell(ColCoveredLines, coverageRow.CoveredLines);
                row.SetCell(ColUncoveredLines, coverageRow.UncoveredLines);
                row.SetCell(ColProjectPathName, coverageRow.ProjectPathName);
                row.SetCell(ColSourcePathName, coverageRow.SourceFilePathName);

                i++;
            }

            sheet.FreezeTopRow();
            ApplyPercentageFormatting(ColCoverage, 1, i);
        }
    }
}
