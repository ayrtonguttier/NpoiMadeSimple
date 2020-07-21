using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;

namespace NpoiMadeSimple.Excel
{
    public class Cell
    {
        private ICell cell;
        private CellReference reference;

        internal Cell(ICell cell)
        {
            this.cell = cell;
            this.reference = new CellReference(cell);
        }

        public override string ToString()
        {
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                            return cell.DateCellValue.ToString();
                        else
                            return cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue.ToString();
                    default:
                        return "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading value from {reference.FormatAsString()}", ex);
            }
        }

    }
}
