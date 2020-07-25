using NpoiMadeSimple.Excel.Typed;
using System;

namespace Test.Model
{
    public class FileModel
    {
        [SheetColumn("OrderDate")]
        public DateTime OrderDate { get; set; }

        [SheetColumn("Region")]
        public string Region { get; set; }

        [SheetColumn("Rep")]
        public string Rep { get; set; }

        [SheetColumn("Item")]
        public string Item { get; set; }

        [SheetColumn("Units")]
        public int Units { get; set; }

        [SheetColumn("Unit Cost")]
        public double UnitCost { get; set; }

        [SheetColumn("Total")]
        public double Total { get; set; }
    }
}
