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
    class Sample02
    {
        public void Run()
        {
            string tempPath = @"模版\PeopleInfo.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, "人员信息");
                List<PeopleInfo> list;
                try
                {
                    list = EpplusHelper.GetList<PeopleInfo>(ws, 2);
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

    internal class PeopleInfo
    {
        public string 名字 { get; set; }
        [EnumUndefined("{0}的性别'{1}'填写不正确","名字","性别" )]
        public Gender? 性别 { get; set; }
        public DateTime? 出生日期 { get; set; }
        public string 身份证号码 { get; set; }
        public int 年龄 { get; set; }
    }

    public enum Gender
    {
        男 = 1,
        女 = 2,
    }

}
