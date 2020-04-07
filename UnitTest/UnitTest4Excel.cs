using System;
using System.Collections.Generic;
using EasyExcel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTest
{
    [TestClass]
    public class UnitTest4Excel
    {
        List<string> ExcelFiles;
        public UnitTest4Excel()
        {
            ExcelFiles = new List<string>();
            ExcelFiles.Add(@".\static\excel.xlsx");
            ExcelFiles.Add(@".\static\excel95.xls");
            ExcelFiles.Add(@".\static\excel97-2004.xls");
        }

        [TestMethod]
        public void TestReadExcel()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var wb = excel.OpenExcel(filePath);
                if (wb != null)
                {
                    Assert.IsTrue(true);
                }
                else
                {
                    Assert.IsTrue(false);
                }

                wb.Close();
            }
        }


        [TestMethod]
        public void TestWriteCell()
        {
            IExcelPlugin excel = new ExcelPlugin();
            var workbook = excel.CreateExcel();
            var sheet = excel.CreateSheet(workbook, "shee1");
            var result = excel.WriteCell(sheet, 0, 1, "name");
            result = excel.WriteCell(sheet, 1, 1, 1);
            result = excel.WriteCell(sheet, 2, 1, 23.90);
            result = excel.WriteCell(sheet, 3, 1, 3999999999.00000003);

            if (result)
            {
                Assert.IsTrue(true);
            }
            else
            {
                Assert.IsTrue(false);
            }

            excel.Save(workbook, "");
        }

        [TestMethod]
        public void TestSetCell2()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                var sheet = excel.GetSheet(workbook, 0);
                if (sheet == null)
                {
                    continue;
                }

                var result = excel.WriteCell(sheet, 5, 5, "just for test value");
                excel.Save(workbook, filePath);

                if (result)
                {
                    Assert.IsTrue(true);
                }
                else
                {
                    Assert.IsTrue(false);
                }
            }
        }

        [TestMethod]
        public void TestReadRow()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                var sheet = excel.GetSheet(workbook, 0);
                var result = excel.ReadRow(sheet, 5);
                if (result.Count >= 0)
                {
                    Assert.IsTrue(true);
                }
                else
                {
                    Assert.IsTrue(false);
                }
            }
        }

        [TestMethod]
        public void TestReadRange()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                var sheet = excel.GetSheet(workbook, 0);
                var result = excel.ReadRange(sheet, "A1:F6");

                foreach (var item in result)
                {
                    Console.WriteLine(item.ToString());
                }

                if (result.Count == 36)
                {
                    Assert.IsTrue(true);
                }
                else
                {
                    Assert.IsTrue(false);
                }
            }
        }

        [TestMethod]
        public void TestGetRowsCount()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                var sheet = excel.GetSheet(workbook, 0);
                var result = excel.GetRowsCount(sheet);
                Console.WriteLine(result);
            }
        }

        [TestMethod]
        public void TestGetColumsCount()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                var sheet = excel.GetSheet(workbook, 0);
                var result = excel.GetColumsCount(sheet);
                Console.WriteLine(result);
            }
        }

        [TestMethod]
        public void TestDeleteSheet()
        {
            IExcelPlugin excel = new ExcelPlugin();

            foreach (var filePath in ExcelFiles)
            {
                var workbook = excel.OpenExcel(filePath);
                workbook = excel.DeleteSheet(workbook, 0);
                excel.Close(workbook, true);
            }
        }
    }
}
