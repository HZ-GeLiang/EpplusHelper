using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using EPPlusExtensions.Attributes;

namespace SampleApp._03读取excel内容
{
    class Sample11
    {
        public void Run()
        {
            string filePath = @"模版\03读取excel内容\Sample11.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                try
                {
                    var args = EPPlusHelper.GetExcelListArgsDefault<UserLeaveStat>(ws, 3);
                    var list = EPPlusHelper.GetList<UserLeaveStat>(args);
                    ObjectDumper.Write(list);
                    Console.WriteLine("读取完毕");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }

            Console.ReadKey();
        }

        class UserLeaveStat
        {
            public string 序号 { get; set; }
            public string 姓名 { get; set; }

            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public string JanuaryStatistics { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public string FebruaryStatistics { get; set; }
        }

    }
}
