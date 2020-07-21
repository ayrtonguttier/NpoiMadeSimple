using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Lifetime;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace NpoiMadeSimple.Excel
{
    public class Sheet
    {
        private Dictionary<string, int> header;
        private int headerRow;
        private readonly ISheet sheet;

        public Row this[int pos]
        {
            get
            {
                if (pos <= headerRow || pos >= sheet.LastRowNum)
                    throw new IndexOutOfRangeException();

                return new Row(header, sheet.GetRow(pos));
            }
        }

        public int Length
        {
            get { return sheet.LastRowNum; }
        }

        public Sheet(ISheet sheet)
        {
            this.header = ReadHeader(sheet, 0);
            this.headerRow = 0;
            this.sheet = sheet;
        }
        public Sheet(ISheet sheet, int headerRowNumber)
        {
            this.header = ReadHeader(sheet, headerRowNumber);
            this.headerRow = headerRowNumber;
            this.sheet = sheet;
        }
        private Dictionary<string, int> ReadHeader(ISheet sheet, int headerRow)
        {
            Dictionary<string, int> cellCount = new Dictionary<string, int>();
            Dictionary<string, int> header = new Dictionary<string, int>();

            var row = sheet.GetRow(headerRow);
            for (int i = row.FirstCellNum; i < row.LastCellNum; i++)
            {
                string cell = new Cell(row.GetCell(i)).ToString();

                if (!cellCount.ContainsKey(cell))
                {
                    cellCount[cell] = 0;
                    header.Add(cell, i);
                }
                else
                {
                    cellCount[cell] = ++cellCount[cell];
                }
            }
            return header;
        }





    }
}
