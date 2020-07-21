using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace NpoiMadeSimple.Excel
{
    public class Row
    {
        private Dictionary<string, int> header;
        private IRow rows;

        public Cell this[string column]
        {
            get
            {
                return new Cell(rows.GetCell(header[column]));
            }
        }


        internal Row(Dictionary<string, int> header, IRow rows)
        {
            this.header = header;
            this.rows = rows;
        }





    }
}
