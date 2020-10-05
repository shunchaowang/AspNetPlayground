using System;
using NPOI.SS.UserModel;

namespace AspNetCore.Controllers.Utils
{
    public static class NpoiExtension
    {
        public static string GetFormattedCellValue(this ICell cell, IFormulaEvaluator evaluator = null)
        {
            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;

                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            // DateTime date = cell.DateCellValue;
                            // ICellStyle style = cell.CellStyle;
                            // // Excel uses lowercase m for month whereas .Net uses uppercase
                            // string format = style.GetDataFormatString().Replace('m', 'M');
                            // return date.ToString(format);

                            try
                            {
                                return cell.DateCellValue.ToString();
                            }
                            catch (NullReferenceException)
                            {
                                return DateTime.FromOADate(cell.NumericCellValue).ToString();
                            }
                        }
                        return cell.NumericCellValue.ToString();

                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";

                    case CellType.Formula:
                        if (evaluator != null)
                        {
                            return GetFormattedCellValue(evaluator.EvaluateInCell(cell));
                        }
                        else
                        {
                            return cell.CellFormula;
                        }

                    case CellType.Error:
                        return FormulaError.ForInt(cell.ErrorCellValue).String;
                }
            }

            // null or blank cell, or unknown cell type
            return string.Empty;
        }
    }
}