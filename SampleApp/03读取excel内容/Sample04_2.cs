using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SampleApp._03读取excel内容
{
    public class Sample04_2
    {
        public static IEnumerable<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample04.xlsx";
            var wsName = "合并行读取";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                args.AddKVSourceByPropName(nameof(args.Model.部门评分), GetSource_部门评分(args.Model));
                var list = EPPlusHelper.GetList(args);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        private static KvSource<long, string> GetSource_部门评分(ExcelModel propModel)
        {
            var source = propModel.部门评分.CreateKVSource();
            source.Add(1, "非常不满意", "very bad");
            source.Add(2, "不满意","bad");
            source.Add(3, "一般","just so so");
            source.Add(4, "满意","good");
            source.Add(5, "非常满意","very good");
            return source;
        }

        public class ExcelModel
        {
            public string 序号 { get; set; }
            public string 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }

            [KVSet("部门评分")]
            public KV<long, string> 部门评分 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门 == y.部门 &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字 &&
                       Helper.GetEquals_KV(this.部门评分, y.部门评分);
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.GetHashCode() +
                       this.部门负责人.GetHashCode() +
                       this.部门负责人确认签字.GetHashCode() +
                       Helper.GetHashCode_KV(this.部门评分);
            }
        }
    }
}
