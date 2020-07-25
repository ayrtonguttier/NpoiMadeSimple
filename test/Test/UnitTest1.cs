using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NpoiMadeSimple.Excel;
using NpoiMadeSimple.Excel.Typed;
using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using Test.Model;

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

            for (int i = 0; i < sheet.Length; i++)
            {
                string resultado = string.Join("|", sheet[i]["OrderDate"], sheet[i]["Region"], sheet[i]["Rep"], sheet[i]["Item"], sheet[i]["Units"], sheet[i]["Unit Cost"], sheet[i]["Total"]);
                System.Diagnostics.Debug.WriteLine(resultado);
            }
        }

        [TestMethod]
        public void TestMethod2()
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

            foreach (var row in sheet)
            {
                string resultado = string.Join("|", row["OrderDate"], row["Region"], row["Rep"], row["Item"], row["Units"], row["Unit Cost"], row["Total"]);
                System.Diagnostics.Debug.WriteLine(resultado);
            }
        }


        [TestMethod]
        public void TestMethod3()
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

            foreach (var row in sheet)
            {
                string resultado = string.Join("|", row.Select(x => x.ToString()).ToArray());
                System.Diagnostics.Debug.WriteLine(resultado);
            }
        }


        [TestMethod]
        public void TestMethod4()
        {
            CultureInfo.CurrentCulture = new CultureInfo("en-US");

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

            foreach (var row in sheet)
            {
                string resultado = string.Join("|", row.Select(x => x.ToString()).ToArray());
                System.Diagnostics.Debug.WriteLine(resultado);
            }
        }

        [TestMethod]
        public void TestMethod5()
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


            Sheet<FileModel> sheet = new Sheet<FileModel>(workbook.GetSheet("SalesOrders"));

            foreach (var item in sheet.AllItens())
            {
                System.Diagnostics.Debug.WriteLine(item.Item);
            }
        }

        [TestMethod]
        public void TestMethod6()
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

            Action action = () =>
            {
                Sheet<FileModelError> sheet = new Sheet<FileModelError>(workbook.GetSheet("SalesOrders"));

                foreach (var item in sheet.AllItens())
                {
                    System.Diagnostics.Debug.WriteLine(item.Item);
                }
            };

            Assert.ThrowsException<Exception>(action);
        }

    }
}
