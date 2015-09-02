using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TestParser.Core
{
    public partial class XLSXWriter : ITestResultWriter
    {
        readonly int yellowBand;
        readonly int greenBand;

        public XLSXWriter(int yellowBand, int greenBand)
        {
            this.yellowBand = yellowBand;
            this.greenBand = greenBand;
        }

        public void WriteResults(Stream s, ParsedData parsedData)
        {
            IWorkbook workbook = new XSSFWorkbook();

            if (parsedData.ResultLines.Count() > 0)
            {
                var trsumWriter = new XLSXTestResultSummarySheetWriter(workbook.CreateSheet("TestResults - Summary"));
                trsumWriter.CreateSheet(yellowBand, greenBand, parsedData);

                var trWriter = new XLSXTestResultSheetWriter(workbook.CreateSheet("TestResults"));
                trWriter.CreateSheet(yellowBand, greenBand, parsedData);
            }

            if (parsedData.NCrunchCoverageData.Count() > 0)
            {
                var ncsumWriter = new XLSXNCrunchCoverageSummarySheetWriter(workbook.CreateSheet("NCrunch Coverage - Summary"));
                ncsumWriter.CreateSheet(yellowBand, greenBand, parsedData.NCrunchCoverageData);

                var ncWriter = new XLSXNCrunchCoverageSheetWriter(workbook.CreateSheet("NCrunch Coverage"));
                ncWriter.CreateSheet(yellowBand, greenBand, parsedData.NCrunchCoverageData);
            }

            workbook.Write(s);
        }
    }
}
