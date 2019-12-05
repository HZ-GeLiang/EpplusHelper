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
    class Sample06
    {
        public void Run()
        {
            string filePath = @"模版\03读取excel内容\Sample06.xlsx";
            var wsName = "Sheet1";
            using( var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var list = EPPlusHelper.GetList<ysbm>(ws, 2);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
            }
        }
        internal class ysbm
        {
            public string 序号 { get; set; }
            [Required(ErrorMessage = "部门不允许为空")]
            public string 部门 { get; set; }

            //[StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
            //public string 部门Id { get; set; }

            [Required(ErrorMessage = "部门Id不能为空")]
            
            [Range(100, 99999, ErrorMessage = "值必须在[101,99999]之间")]
            public long 部门Id { get; set; }

            public string 预算部门 { get; set; }
            public string 预算部门负责人 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }
        }
    }
}
