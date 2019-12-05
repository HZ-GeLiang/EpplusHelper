using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp._03读取excel内容
{
    class Sample08
    {
        public void Run()
        {
            string filePath = @"模版\03读取excel内容\Sample08.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<Test02_3>(ws, 2);
                args.POCO_Property_AutoRename_WhenRepeat = true;
                args.POCO_Property_AutoRenameFirtName_WhenRepeat = false;
                var list = EPPlusHelper.GetList<Test02_3>(args);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
            }
        }
        internal class Test02_3
        {
            public string 名字 { get; set; }
            public string 名字2 { get; set; }
            public string 名字3 { get; set; }
        }
    }
}
