﻿using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Lifetime;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace NpoiMadeSimple.Excel
{
    public class Sheet : IEnumerable<Row>
    {
        private Dictionary<string, int> header;
        private int headerRow;
        private readonly ISheet sheet;

        public Row this[int pos]
        {
            get
            {
                int lastPosition = sheet.LastRowNum - (headerRow + 1);
                int position = pos + (headerRow + 1);

                if (pos < 0 || pos > lastPosition)
                    throw new IndexOutOfRangeException();

                return new Row(header, sheet.GetRow(position));
            }
        }

        public int Length
        {
            get { return sheet.LastRowNum - headerRow; }
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
                    header.Add(string.Concat(cell, "_", ++cellCount[cell]), i);
                }
            }
            return header;
        }

        public IEnumerator<Row> GetEnumerator()
        {
            return new RowEnumerator(sheet, header, headerRow);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private class RowEnumerator : IEnumerator<Row>
        {
            private ISheet sheet;
            private Dictionary<string, int> header;
            private bool isDisposed;
            private readonly int firstRow;
            private int position;

            public Row Current { get; private set; }

            object IEnumerator.Current => Current;

            public RowEnumerator(ISheet sheet, Dictionary<string, int> header, int headerRow)
            {
                this.sheet = sheet;
                this.header = header;
                this.firstRow = headerRow + 1;
            }

            public void Dispose()
            {
                if (isDisposed)
                    return;

                sheet = null;
                header = null;

                isDisposed = true;
            }

            public bool MoveNext()
            {
                if (++position > sheet.LastRowNum)
                    return false;

                Current = new Row(header, sheet.GetRow(position));
                return true;
            }

            public void Reset()
            {
                position = firstRow;
            }
        }
    }
}
