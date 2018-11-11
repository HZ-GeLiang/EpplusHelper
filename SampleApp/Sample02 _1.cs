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
    class Sample02_1
    {
        public void Run()
        {
            string tempPath = @"模版\dept.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                List<ysbm> list;
                try
                {
                    list = EpplusHelper.GetList<ysbm>(ws, 2);
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

    internal class ysbm
    {
        public string 序号 { get; set; }
        public string 部门 { get; set; }
        public string 预算部门 { get; set; }
        public string 预算部门负责人 { get; set; }
        public string 部门负责人 { get; set; }
        public string 部门负责人确认签字 { get; set; }


    }


}
