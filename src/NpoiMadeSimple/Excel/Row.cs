using NPOI.SS.UserModel;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace NpoiMadeSimple.Excel
{
    public class Row : IEnumerable<Cell>
    {
        private Dictionary<string, int> header;
        private IRow row;

        public Cell this[string column]
        {
            get
            {
                return new Cell(row.GetCell(header[column]));
            }
        }

        internal Row(Dictionary<string, int> header, IRow rows)
        {
            this.header = header;
            this.row = rows;
        }

        public IEnumerator<Cell> GetEnumerator()
        {
            return new CellEnumerator(row);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private class CellEnumerator : IEnumerator<Cell>
        {
            private IRow row;
            private int position;
            public Cell Current { get; private set; }

            private bool isDisposed;

            object IEnumerator.Current => Current;

            public CellEnumerator(IRow row)
            {
                this.row = row;
                position = row.FirstCellNum - 1;
            }

            public void Dispose()
            {
                if (isDisposed)
                    return;

                row = null;
                Current = null;
                isDisposed = true;
            }

            public bool MoveNext()
            {
                if (++position >= row.LastCellNum)
                    return false;

                Current = new Cell(row.GetCell(position));
                return true;
            }

            public void Reset()
            {
                position = row.FirstCellNum;
            }
        }
    }
}
