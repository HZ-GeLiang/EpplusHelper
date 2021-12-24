using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample12
    {
        public static void Run()
        {
            Run(true);
        }
        public static void Run(bool OnlyShowColomn)
        {
            Run<ExcelModel2>(OnlyShowColomn);
        }
        public static List<T> Run<T>(bool OnlyShowColomn) where T : class, new()
        {
            List<T> excelList = null;
            var errorMsg = EPPlusHelper.GetListErrorMsg(() =>
            {
                string filePath = @"模版\03读取excel内容\Sample12.xlsx";
                using (var fs = EPPlusHelper.GetFileStream(filePath))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<T>(ws, 3);
                    args.GetList_NeedAllException = true; //默认false

                    args.GetList_ErrorMessage_OnlyShowColomn = OnlyShowColomn; //默认false
                    //true:  => 参数名: 姓名(B列)
                    //false: => 参数名: 姓名(B3,B4,B5)

                    excelList = EPPlusHelper.GetList(args).ToList();
                    ObjectDumper.Write(excelList);
                    Console.WriteLine("读取完毕");
                }
            });

            if (errorMsg?.Length > 0)
            {
                Console.WriteLine(errorMsg);
                throw new Exception(errorMsg);
            }
            else
            {
                return excelList;
            }

        }

        public class ExcelModelBase
        {
            public int 序号 { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public int JanuaryStatistics { get; set; }

            [ExcelColumnIndex(5)]
            [DisplayExcelColumnName("请假次数")]
            public int FebruaryStatistics { get; set; }
        }

        public class ExcelModel1 : ExcelModelBase
        {
            public string 姓名 { get; set; }
            public string 班级 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel1 y = (ExcelModel1)obj;

                return this.序号 == y.序号 &&
                       this.姓名 == y.姓名 &&
                       this.班级 == y.班级 &&
                       this.JanuaryStatistics == y.JanuaryStatistics &&
                       this.FebruaryStatistics == y.FebruaryStatistics;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.姓名.GetHashCode() +
                       this.班级.GetHashCode() +
                       this.JanuaryStatistics.GetHashCode() +
                       this.FebruaryStatistics.GetHashCode();
            }
        }

        /// <summary>
        /// 单元测试用的
        /// </summary>
        public class ExcelModel2 : ExcelModelBase
        {
            public int 姓名 { get; set; }
            public int 班级 { get; set; }
            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel2 y = (ExcelModel2)obj;

                return this.序号 == y.序号 &&
                       this.姓名 == y.姓名 &&
                       this.班级 == y.班级 &&
                       this.JanuaryStatistics == y.JanuaryStatistics &&
                       this.FebruaryStatistics == y.FebruaryStatistics;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.姓名.GetHashCode() +
                       this.班级.GetHashCode() +
                       this.JanuaryStatistics.GetHashCode() +
                       this.FebruaryStatistics.GetHashCode();
            }
        }
    }
}
