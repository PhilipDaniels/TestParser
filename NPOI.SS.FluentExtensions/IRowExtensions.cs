using System;
using NPOI.SS.UserModel;

namespace NPOI.SS.FluentExtensions
{
    /// <summary>
    /// Extensions to the NPOI IRow class.
    /// </summary>
    public static class IRowExtensions
    {
        /// <summary>
        /// Creates a cell and sets it to a boolean value.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetCell(this IRow row, int column, bool value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets it to a DateTime value.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetCell(this IRow row, int column, DateTime value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets it to a double value.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetCell(this IRow row, int column, double value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets it to a rich text value.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetCell(this IRow row, int column, IRichTextString value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets it to a string value.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetCell(this IRow row, int column, string value)
        {
            var cell = row.CreateCell(column);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets its formula.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="formula">The formula to set.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetFormula(this IRow row, int column, string formula)
        {
            var cell = row.CreateCell(column);
            cell.SetCellFormula(formula);
            return cell;
        }

        /// <summary>
        /// Creates a cell and sets its formula.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="formula">The formula to set.</param>
        /// <param name="prms">Optional parameters to substitute into the formula string.</param>
        /// <returns>The newly created cell.</returns>
        public static ICell SetFormula(this IRow row, int column, string formula, params object[] prms)
        {
            if (prms != null && prms.Length > 0)
                formula = String.Format(formula, prms);
            return SetFormula(row, column, formula);
        }
    }
}
