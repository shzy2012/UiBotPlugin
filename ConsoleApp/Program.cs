using EasyExcel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            IExcelPlugin excel = new ExcelPlugin();
            var workbook = excel.CreateExcel();
            var sheet = excel.CreateSheet(workbook, "sheet1");

            //测试1

            string json = @"['Small','Medium','Large']";
            JArray data = JArray.Parse(json);
            var ok = excel.WriteRow(sheet, 0, 1, data);

            //测试2

            var sheet2 = excel.CreateSheet(workbook, "sheet2");

            json = @"[['a1','b1','c1'],['a2','b2','c2']]";
            data = JArray.Parse(json);
            ok = excel.WriteRange(sheet2, 0, 0, data);

            ok = excel.Save(workbook);
            Console.WriteLine("保存结果:{0}", ok);

            Console.ReadLine();
        }
    }
}
