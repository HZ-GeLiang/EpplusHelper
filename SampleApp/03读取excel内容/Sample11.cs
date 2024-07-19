using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample11
    {
        //列名重复: 指定实体属性
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample11.xlsx";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 3);
                var list = EPPlusHelper.GetList(args).ToList();
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
            public string 姓名 { get; set; }

            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public int JanuaryStatistics { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public int FebruaryStatistics { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
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
    }
}
