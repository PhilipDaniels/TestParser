using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.Util;
using NPOI.SS.FluentExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using TestParser.Core.XL;

namespace TestParser.Core
{
    public class XLSXTestResultWriter : ITestResultWriter
    {
        IWorkbook workbook;
        ISheet summarySheet;
        ISheet resultsSheet;

        public void WriteResults(Stream s, IEnumerable<TestResult> testResults)
        {
            workbook = new XSSFWorkbook();
            summarySheet = workbook.CreateSheet("Summary");
            resultsSheet = workbook.CreateSheet("TestResults");

            CreateResultsSheet(testResults);
            CreateSummarySheet(testResults);

            workbook.Write(s);
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
                    if (oc.Outcome == ResultOutcomeSummary.FailedOutcome && oc.NumTests > 0)
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

            return rowNum;
        }

        void CreateSummarySheet(IEnumerable<TestResult> testResults)
        {
            var sba = TestResultSummary.SummariseByAssembly(testResults);
            int rowNum = CreateSummary("Summary By Assembly", 0, sba);

            var sbc = TestResultSummary.SummariseByClass(testResults);
            rowNum++;
            rowNum = CreateSummary("Summary By Class", rowNum, sbc);

            // Set the widths.
            summarySheet.SetColumnWidth(0, 10000);
            summarySheet.SetColumnWidth(1, 10000);
            summarySheet.SetColumnWidth(2, 3000);
            summarySheet.SetColumnWidth(3, 2500);
            for (int colNum = 4; colNum <= 4 + sba.ElementAt(0).Outcomes.Count(); colNum++)
                summarySheet.SetColumnWidth(colNum, 2500);
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
            const int ColDuration = 9;
            const int ColResultsPathName = 10;
            const int ColResultsFileName = 11;
            const int ColAssemblyPathName = 12;
            const int ColFullClassName = 13;
            const int ColComputerName = 14;
            

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
            row.SetCell(ColDuration, "Duration").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColDurationInSeconds, "DurationInSeconds").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColOutcome, "Outcome").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColErrorMessage, "ErrorMessage").HeaderStyle().ApplyStyle(); ;
            row.SetCell(ColStackTrace, "StackTrace").HeaderStyle().ApplyStyle(); ;

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
                row.SetCell(ColStartTime, r.StartTime).FormatLongDate().ApplyStyle();
                row.SetCell(ColEndTime, r.StartTime).FormatLongDate().ApplyStyle();
                row.SetCell(ColDuration, r.Duration.ToString("c")).Alignment(HorizontalAlignment.Right).ApplyStyle();
                row.SetCell(ColDurationInSeconds, r.DurationInSeconds).Format("0.00").ApplyStyle();
                row.SetCell(ColOutcome, r.Outcome);
                row.SetCell(ColErrorMessage, r.ErrorMessage);
                row.SetCell(ColStackTrace, r.StackTrace);

                i++;
            }

            // Freeze the header row.
            resultsSheet.CreateFreezePane(0, 1, 0, 1);

            // Set the widths.
            resultsSheet.SetColumnWidth(0, 10000);
            resultsSheet.SetColumnWidth(1, 10000);
            resultsSheet.SetColumnWidth(2, 10000);
            resultsSheet.SetColumnWidth(3, 3000);
            resultsSheet.SetColumnWidth(4, 3000);
            resultsSheet.SetColumnWidth(5, 3000);
            resultsSheet.SetColumnWidth(6, 3000);
            resultsSheet.SetColumnWidth(7, 5500);
            resultsSheet.SetColumnWidth(8, 5500);
            resultsSheet.SetColumnWidth(9, 5500);
        }

        /*
        ICell SetPercent(IRow row, int column, double number)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(number);
            if (number < 0.9)
                cell.CellStyle = PercentRedStyle;
            else if (number < 1.0)
                cell.CellStyle = PercentYellowStyle;
            else
                cell.CellStyle = PercentGreenStyle;

            return cell;
        }

        ICellStyle PercentGreenStyle
        {
            get
            {
                if (percentGreenStyle == null)
                {
                    percentGreenStyle = workbook.CreateCellStyle();
                    percentGreenStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
                    percentGreenStyle.FillPattern = FillPattern.SolidForeground;
                    percentGreenStyle.DataFormat = PercentFormat;
                }

                return percentGreenStyle;
            }
        }
        ICellStyle percentGreenStyle;

        ICellStyle PercentYellowStyle
        {
            get
            {
                if (percentYellowStyle == null)
                {
                    percentYellowStyle = workbook.CreateCellStyle();
                    percentYellowStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                    percentYellowStyle.FillPattern = FillPattern.SolidForeground;
                    percentYellowStyle.DataFormat = PercentFormat;
                }

                return percentYellowStyle;
            }
        }
        ICellStyle percentYellowStyle;

        ICellStyle PercentRedStyle
        {
            get
            {
                if (percentRedStyle == null)
                {
                    percentRedStyle = workbook.CreateCellStyle();
                    percentRedStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                    percentRedStyle.FillPattern = FillPattern.SolidForeground;
                    percentRedStyle.DataFormat = PercentFormat;
                }

                return percentRedStyle;
            }
        }
        ICellStyle percentRedStyle;
        */
    }
}
