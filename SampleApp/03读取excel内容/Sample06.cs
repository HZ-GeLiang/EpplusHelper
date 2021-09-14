using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample06
    {
        public static void Run()
        {
            Sample06.Run<Sample06.ExcelModel>();
        }

        public static List<T> Run<T>() where T : class, new()
        {
            string filePath = @"模版\03读取excel内容\Sample06.xlsx";
            var wsName = "Sheet1";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var list = EPPlusHelper.GetList<T>(ws, 2).ToList();
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
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

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门 == y.部门 &&
                       this.部门Id == y.部门Id &&
                       this.预算部门 == y.预算部门 &&
                       this.预算部门负责人 == y.预算部门负责人 &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.GetHashCode() +
                       this.部门Id.GetHashCode() +
                       this.预算部门.GetHashCode() +
                       this.预算部门负责人.GetHashCode() +
                       this.部门负责人.GetHashCode() +
                       this.部门负责人确认签字.GetHashCode();
            }
        }

        public class ExcelModel2
        {
            public int 序号 { get; set; }
            [Required(ErrorMessage = "部门不允许为空")]
            public string 部门 { get; set; }

            //[StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
            //public string 部门Id { get; set; }

            [Required(ErrorMessage = "部门Id不能为空")]

            [Range(101, 99999, ErrorMessage = "值必须在[101,99999]之间")]
            public long 部门Id { get; set; }

            public string 预算部门 { get; set; }
            public string 预算部门负责人 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门 == y.部门 &&
                       this.部门Id == y.部门Id &&
                       this.预算部门 == y.预算部门 &&
                       this.预算部门负责人 == y.预算部门负责人 &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.GetHashCode() +
                       this.部门Id.GetHashCode() +
                       this.预算部门.GetHashCode() +
                       this.预算部门负责人.GetHashCode() +
                       this.部门负责人.GetHashCode() +
                       this.部门负责人确认签字.GetHashCode();
            }
        }

        public class ExcelModel3
        {
            public int 序号 { get; set; }
            [StringLength(10, ErrorMessage = "部门名字长度要在9-10之间", MinimumLength = 9)]
            [Required(ErrorMessage = "部门不允许为空")]
            public string 部门 { get; set; }

            //[StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
            //public string 部门Id { get; set; }

            [Required(ErrorMessage = "部门Id不能为空")]

            [Range(101, 99999, ErrorMessage = "值必须在[101,99999]之间")]
            public long 部门Id { get; set; }

            public string 预算部门 { get; set; }
            public string 预算部门负责人 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门 == y.部门 &&
                       this.部门Id == y.部门Id &&
                       this.预算部门 == y.预算部门 &&
                       this.预算部门负责人 == y.预算部门负责人 &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.GetHashCode() +
                       this.部门Id.GetHashCode() +
                       this.预算部门.GetHashCode() +
                       this.预算部门负责人.GetHashCode() +
                       this.部门负责人.GetHashCode() +
                       this.部门负责人确认签字.GetHashCode();
            }
        }
    }
}
