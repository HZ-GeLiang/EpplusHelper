using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp._03读取excel内容
{
    class Sample01
    {
        public void Run()
        {
            string filePath = @"模版\03读取excel内容\Sample01.xlsx";
            var wsName = "逐行读取";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                try
                {
                    var args = EPPlusHelper.GetExcelListArgsDefault<ysbm>(ws, 2);
                    args.ScanLine = ScanLine.SingleLine;
                    var list = EPPlusHelper.GetList<ysbm>(args);//excel的5,6行是合并的,用SingleLine读取,list的第4,5条数据内容是一样的
                    ObjectDumper.Write(list);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                Console.WriteLine("读取完毕");
            }
        }

        class ysbm
        {
            public string 序号 { get; set; }
            public string 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }
        }
    }
}
