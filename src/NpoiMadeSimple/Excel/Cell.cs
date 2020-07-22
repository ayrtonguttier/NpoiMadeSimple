using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Globalization;

namespace NpoiMadeSimple.Excel
{
    public class Cell
    {
        private ICell cell;
        internal Cell(ICell cell)
        {
            this.cell = cell;
        }

        public void SetValue(object value)
        {
            if (value is bool)
                cell.SetCellValue((bool)value);
            else if (value is DateTime)
                cell.SetCellValue((DateTime)value);
            else if (value is double)
                cell.SetCellValue((double)value);
            else if (value is IRichTextString)
                cell.SetCellValue((IRichTextString)value);
            else if (value is string)
                cell.SetCellValue((string)value);
            else
                throw new ArgumentException($"Invalid value type {value.GetType().Name}.");
        }

        public string GetValue()
        {
            try
            {
                CellType type = cell.CellType;
                if (type == CellType.Formula)
                    type = cell.CachedFormulaResultType;

                switch (type)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString(CultureInfo.CurrentCulture);
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue.ToString(CultureInfo.CurrentCulture);
                        else
                            return cell.NumericCellValue.ToString(CultureInfo.CurrentCulture);
                    case CellType.String:
                        return cell.StringCellValue.ToString();
                    default:
                        return "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading value from {cell.Address}", ex);
            }
        }

        public override string ToString()
        {
            return GetValue();
        }
    }
}
