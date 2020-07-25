using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NpoiMadeSimple.Excel.Typed
{
    [AttributeUsage(AttributeTargets.Property)]
    public class SheetColumnAttribute : Attribute
    {
        private readonly string name;

        public SheetColumnAttribute(string name)
        {
            this.name = name;
        }


        public string GetName()
        {
            return name;
        }
    }
}
