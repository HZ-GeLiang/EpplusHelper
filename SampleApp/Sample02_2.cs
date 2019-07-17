using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using EPPlusExtensions;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_2
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_2.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                var list=  EPPlusHelper.GetList<ysbm>(ws, 2);
                Console.WriteLine("读取完毕");
            }
        }
        internal class ysbm
        {
            public string 序号 { get; set; }
            [Required(ErrorMessage = "部门不允许为空")]
            public string 部门 { get; set; }

            [Required(ErrorMessage = "部门Id不能为空")]
            [StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
            [Range(101, 99999, ErrorMessage = "值必须在[101,99999]之间")]
            //public string 部门Id { get; set; }
            public long 部门Id { get; set; }

            public string 预算部门 { get; set; }
            public string 预算部门负责人 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }
        }
    }
}
