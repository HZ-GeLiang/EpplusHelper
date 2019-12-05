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
    class Sample07
    {
        public void Run()
        {

            string filePath = @"模版\03读取excel内容\Sample07.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var list = EPPlusHelper.GetList<PeopleInfo>(ws, 2);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
            }
        }
        class PeopleInfo
        {
            public string 序号 { get; set; }
            //[Unique()]
            public string 名字 { get; set; }
            [EnumUndefined("{0}的性别'{1}'填写不正确", "名字", "性别")]
            public Gender? 性别 { get; set; }
            public DateTime? 出生日期 { get; set; }
            public string 身份证号码 { get; set; }
            public int 年龄 { get; set; }
        }
        enum Gender
        {
            男 = 1,
            女 = 2,
            未知 = 3,
        }
    }
}
