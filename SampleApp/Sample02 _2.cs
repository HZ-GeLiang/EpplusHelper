using EpplusExtensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using EpplusExtensions.Attributes;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_2
    {
        public void Run()
        {
            string tempPath = @"模版\dept_02_2.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                List<ysbm2> list;
                try
                {
                    list = EpplusHelper.GetList<ysbm2>(ws, 2);
                }
                catch (Exception e)
                {
                    if (e.Message.Contains("类型中没有定义该属性"))
                    {
                        StringBuilder excelFileds = new StringBuilder();
                        foreach (var item in e.Message.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries))
                        {
                            int start = item.IndexOf("'值'");
                            int end = item.IndexOf("'在'");
                            string excelFiled = item.Substring(start + 3, end - start - 3);
                            excelFileds.Append(excelFiled).Append(",");
                        }
                        excelFileds.RemoveLastChar(',');
                        throw new Exception("提供了excel模版之外列:" + excelFileds);
                    }
                    else
                    {
                        throw e;
                    }
                }

                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class ysbm2
    {
        public string 序号 { get; set; }
        [System.ComponentModel.DataAnnotations.Required(ErrorMessage = "部门不允许为空")]
        public string 部门 { get; set; }

        [System.ComponentModel.DataAnnotations.Required(ErrorMessage = "部门Id不能为空")]
        [System.ComponentModel.DataAnnotations.StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
        [System.ComponentModel.DataAnnotations.Range(101, 99999, ErrorMessage = "值必须在[101,99999]之间")]
        //public string 部门Id { get; set; }
        public long 部门Id { get; set; }

        public string 预算部门 { get; set; }
        public string 预算部门负责人 { get; set; }
        public string 部门负责人 { get; set; }
        public string 部门负责人确认签字 { get; set; }
    }
}
