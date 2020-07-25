using NpoiMadeSimple.Excel.Typed;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.Model
{
    public class FileModelError
    {
        [SheetColumn("OrderDate")]
        public DateTime OrderDate { get; set; }

        [SheetColumn("Region")]
        public string Region { get; set; }

        [SheetColumn("Rep")]
        public int Rep { get; set; }

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

