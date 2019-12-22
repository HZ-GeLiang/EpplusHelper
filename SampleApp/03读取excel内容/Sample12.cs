using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using EPPlusExtensions.Attributes;

namespace SampleApp._03读取excel内容
{
    public class Sample12
    {
        public static void Run()
        {
            Run<ExcelModel2>();
        }
        public static List<T> Run<T>() where T : class, new()
        {
            List<T> list = new List<T>();
            var errorMsg = EPPlusHelper.GetListErrorMsg(() =>
            {
                string filePath = @"模版\03读取excel内容\Sample12.xlsx";
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<T>(ws, 3);
                    args.GetList_NeedAllException = true; //默认false
                    args.GetList_ErrorMessage_OnlyShowColomn = true; //默认false
                                                                     //true:  => 参数名: 姓名(B列)
                                                                     //false: => 参数名: 姓名(B3,B4,B5)

                    list = EPPlusHelper.GetList(args);
                    ObjectDumper.Write(list);
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
                return list;
            }

        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
            public string 姓名 { get; set; }
            //public int 姓名 { get; set; } //要效果,取消注释

            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public int JanuaryStatistics { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public int FebruaryStatistics { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.姓名 == y.姓名 &&
                       this.JanuaryStatistics == y.JanuaryStatistics &&
                       this.FebruaryStatistics == y.FebruaryStatistics;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.姓名.GetHashCode() +
                       this.JanuaryStatistics.GetHashCode() +
                       this.FebruaryStatistics.GetHashCode();
            }
        }

        public class ExcelModel2
        {
            public string 序号 { get; set; }
            //public string 姓名 { get; set; }
            public int 姓名 { get; set; } //要效果,取消注释

            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public string JanuaryStatistics { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public string FebruaryStatistics { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel2 y = (ExcelModel2)obj;

                return this.序号 == y.序号 &&
                       this.姓名 == y.姓名 &&
                       this.JanuaryStatistics == y.JanuaryStatistics &&
                       this.FebruaryStatistics == y.FebruaryStatistics;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.姓名.GetHashCode() +
                       this.JanuaryStatistics.GetHashCode() +
                       this.FebruaryStatistics.GetHashCode();
            }
        }
    }
}
