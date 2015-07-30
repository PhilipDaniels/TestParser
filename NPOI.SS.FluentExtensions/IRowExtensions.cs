using System;
using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    public static class IRowExtensions
    {
        public static ICell SetCell(this IRow row, int column, bool value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCell(this IRow row, int column, DateTime value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCell(this IRow row, int column, double value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCell(this IRow row, int column, IRichTextString value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetCell(this IRow row, int column, string value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        public static ICell SetFormula(this IRow row, int column, string formula)
        {
            var cell = row.CreateCell(column);
            cell.SetCellFormula(formula);
            return cell;
        }

        public static ICell SetFormula(this IRow row, int column, string formula, params object[] prms)
        {
            if (prms != null && prms.Length > 0)
                formula = String.Format(formula, prms);
            return SetFormula(row, column, formula);
        }
    }
}
