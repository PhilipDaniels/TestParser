using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestParser.Core.XL;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;

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

        void CreateSummarySheet(IEnumerable<TestResult> testResults)
        {
            int rowNum = 0;
            int colNum = 0;
            IRow row = summarySheet.CreateRow(rowNum++);


            // =============================== Part 1, "Summary By Assembly" ===============================
            var sba = TestResultSummaryByAssembly.Summarise(testResults);

            SetText(row, 0, "Summary By Assembly").LargeHeaderStyle().ApplyStyle();
            row = summarySheet.CreateRow(rowNum++);

            SetText(row, 0, "Assembly").HeaderStyle().ApplyStyle();
            SetText(row, 1, "Class").HeaderStyle().ApplyStyle();
            SetText(row, 2, "Time (secs)").HeaderStyle().ApplyStyle();
            SetText(row, 3, "Percent").HeaderStyle().ApplyStyle();
            SetText(row, 4, "Total").HeaderStyle().ApplyStyle();
            colNum = 5;
            foreach (var oc in sba.ElementAt(0).Outcomes)
            {
                SetText(row, colNum++, oc.Outcome).HeaderStyle().ApplyStyle();
            }

            foreach (var s in sba)
            {
                row = summarySheet.CreateRow(rowNum++);
                SetText(row, 0, s.AssemblyFileName);
                SetText(row, 1, "n.a.").Alignment(HorizontalAlignment.CenterSelection).ApplyStyle();
                SetNumber(row, 2, s.TotalDurationInSeconds);
                SetFormula(row, 3, String.Format("F{0}/E{0}", rowNum)).FormatPercentage().ApplyStyle();
                SetNumber(row, 4, s.TotalTests);

                colNum = 5;
                foreach (var oc in s.Outcomes)
                {
                    SetNumber(row, colNum++, oc.NumTests);
                }
            }


            /*
            row = summarySheet.CreateRow(rowNum++);
            SetFormula(row, 1, String.Format("SUM(B3:B{0})", rowNum - 1));
            SetFormula(row, 2, String.Format("", rowNum - 1));
            SetFormula(row, 3, String.Format("SUM(D3:D{0})", rowNum - 1));
            int maxCol = 4 + outcomes.Count();
            for (int cn = 4; cn < maxCol; cn++)
            {
                string colRef = CellReference.ConvertNumToColString(cn);
                SetFormula(row, cn, String.Format("SUM({0}3:{0}{1})", colRef, rowNum - 1));
            }

            // =============================== Part 2, "Summary By Assembly" ===============================
            // =============================== Part 3, "Failing Tests" ===============================
            //SetLargeHeader(row, 0, "Summary By Class");
            */
        }


        void CreateResultsSheet(IEnumerable<TestResult> testResults)
        {
            const int ColResultsPathName = 0;
            const int ColResultsFileName = 1;
            const int ColAssemblyPathName = 2;
            const int ColAssemblyFileName = 3;
            const int ColFullClassName = 4;
            const int ColClassName = 5;
            const int ColTestName = 6;
            const int ColComputerName = 7;
            const int ColStartTime = 8;
            const int ColEndTime = 9;
            const int ColDuration = 10;
            const int ColDurationInSeconds = 11;
            const int ColOutcome = 12;
            const int ColErrorMessage = 13;
            const int ColStackTrace = 14;

            IRow row = resultsSheet.CreateRow(0);
            //SetHeader(row, ColResultsPathName, "ResultsPathName");
            //SetHeader(row, ColResultsFileName, "ResultsFileName");
            //SetHeader(row, ColAssemblyPathName, "AssemblyPathName");
            //SetHeader(row, ColAssemblyFileName, "AssemblyFileName");
            //SetHeader(row, ColFullClassName, "FullClassName");
            //SetHeader(row, ColClassName, "ClassName");
            //SetHeader(row, ColTestName, "TestName");
            //SetHeader(row, ColComputerName, "ComputerName");
            //SetHeader(row, ColStartTime, "StartTime");
            //SetHeader(row, ColEndTime, "EndTime");
            //SetHeader(row, ColDuration, "Duration");
            //SetHeader(row, ColDurationInSeconds, "DurationInSeconds");
            //SetHeader(row, ColOutcome, "Outcome");
            //SetHeader(row, ColErrorMessage, "ErrorMessage");
            //SetHeader(row, ColStackTrace, "StackTrace");

            int i = 1;
            foreach (var r in testResults)
            {
                row = resultsSheet.CreateRow(i);
                SetText(row, ColResultsPathName, r.ResultsPathName);
                SetText(row, ColResultsFileName, r.ResultsFileName);
                SetText(row, ColAssemblyPathName, r.AssemblyPathName);
                SetText(row, ColAssemblyFileName, r.AssemblyFileName);
                SetText(row, ColFullClassName, r.FullClassName);
                SetText(row, ColClassName, r.ClassName);
                SetText(row, ColTestName, r.TestName);
                SetText(row, ColComputerName, r.ComputerName);
                SetDate(row, ColStartTime, r.StartTime);
                SetDate(row, ColEndTime, r.StartTime);
                SetText(row, ColDuration, r.Duration.ToString("c"));
                SetNumber(row, ColDurationInSeconds, r.DurationInSeconds);
                SetText(row, ColOutcome, r.Outcome);
                SetText(row, ColErrorMessage, r.ErrorMessage);
                SetText(row, ColStackTrace, r.StackTrace);

                i++;
            }

            // Freeze the header row.
            resultsSheet.CreateFreezePane(0, 1, 0, 1);
        }

        ICell SetText(IRow row, int column, string text)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(text);
            return cell;
        }

        ICell SetDate(IRow row, int column, DateTime date)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(date);
            cell.CellStyle = DateStyle;
            return cell;
        }

        ICell SetNumber(IRow row, int column, double number)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(number);
            return cell;
        }

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



        ICell SetFormula(IRow row, int column, string formula)
        {
            var cell = row.CreateCell(column);
            cell.SetCellFormula(formula);
            return cell;
        }



        ICellStyle DateStyle
        {
            get
            {
                if (dateStyle == null)
                {
                    dateStyle = workbook.CreateCellStyle();
                    IDataFormat dataFormat = workbook.CreateDataFormat();
                    dateStyle.DataFormat = dataFormat.GetFormat("yyyy-MM-dd HH:mm:ss");
                }

                return dateStyle;
            }
        }
        ICellStyle dateStyle;

        short PercentFormat
        {
            get
            {
                if (percentFormat == null)
                { 
                    percentFormat = workbook.CreateDataFormat().GetFormat("0.00%");
                }

                return percentFormat.Value;
            }
        }
        short? percentFormat;

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

        ICellStyle MakeStyle(int backgroundColor, string dataFormat, bool bold, bool italic, int fontHeight)
        {
            return null;
        }
    }
}
