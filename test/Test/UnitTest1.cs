using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NpoiMadeSimple.Excel;
using System.IO;

namespace Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string file = @".\Files\SampleData.xlsx";
            IWorkbook workbook;
            using (FileStream stream = new FileStream(file, FileMode.Open))
            {
                if (file.EndsWith(".xlsx"))
                    workbook = new XSSFWorkbook(stream);
                else
                    workbook = new HSSFWorkbook(stream);
            }

            Sheet sheet = new Sheet(workbook.GetSheet("SalesOrders"));

            for (int i = 0; i< sheet.Length; i++)
            {
                string resultado = string.Join("|", sheet[i]["OrderDate"], sheet[i]["Region"], sheet[i]["Rep"], sheet[i]["Item"], sheet[i]["Units"], sheet[i]["Unit Cost"], sheet[i]["Total"]);
                System.Diagnostics.Debug.WriteLine(resultado);
            }
        }
    }
}
