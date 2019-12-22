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
    public class Sample07
    {
        public static void Run()
        {
            return;//下面代码肯定异常
            string filePath = @"模版\03读取excel内容\Sample07.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                var list = EPPlusHelper.GetList(args);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
            }
        }
        public static List<T> Run<T>() where T : class, new()
        {

            string filePath = @"模版\03读取excel内容\Sample07.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var list = EPPlusHelper.GetList<T>(ws, 2);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }
        public class ExcelModel
        {
            public int 序号 { get; set; }
            //[Unique()]
            public string 名字 { get; set; }
            // [EnumUndefined("{0}的性别'{1}'填写不正确", "名字", "性别")]
            public Gender? 性别 { get; set; }
            public DateTime? 出生日期 { get; set; }
            public string 身份证号码 { get; set; }
            public int 年龄 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.名字 == y.名字 &&
                       this.性别 == y.性别 &&
                       this.出生日期 == y.出生日期 &&
                       this.身份证号码 == y.身份证号码 &&
                       this.年龄 == y.年龄;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.名字.GetHashCode() +
                       this.性别.GetHashCode() +
                       this.出生日期.GetHashCode() +
                       this.身份证号码.GetHashCode() +
                       this.年龄.GetHashCode();
            }
        }

        public class ExcelModel2
        {
            public int 序号 { get; set; }
            //[Unique()]
            public string 名字 { get; set; }
            [EnumUndefined("{0}的性别'{1}'填写不正确", "名字", "性别")]
            public Gender? 性别 { get; set; }
            public DateTime? 出生日期 { get; set; }
            public string 身份证号码 { get; set; }
            public int 年龄 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.名字 == y.名字 &&
                       this.性别 == y.性别 &&
                       this.出生日期 == y.出生日期 &&
                       this.身份证号码 == y.身份证号码 &&
                       this.年龄 == y.年龄;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.名字.GetHashCode() +
                       this.性别.GetHashCode() +
                       this.出生日期.GetHashCode() +
                       this.身份证号码.GetHashCode() +
                       this.年龄.GetHashCode();
            }
        }
        public enum Gender
        {
            男 = 1,
            女 = 2,
            未知 = 3,
        }
    }
}
