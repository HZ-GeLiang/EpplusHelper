using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EPPlusExtensions.CustomModelType;

namespace SampleApp._03读取excel内容
{
    public class Sample05
    {
        public static List<ExcelModel> Run()
        {
            var dataSource = new Dictionary<string, long?>();
            dataSource.Add("事业1部", 1);
            dataSource.Add("事业2部", 2);
            dataSource.Add("事业3部", null);
            return Run(dataSource);
        }

        //数据源中有字典数据

        public static List<ExcelModel> Run(Dictionary<string, long?> dataSource)
        {
            string filePath = @"模版\03读取excel内容\Sample05.xlsx";
            var wsName = 1;
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                args.ScanLine = ScanLine.SingleLine;

                //var dataSource = propModel.部门2.CreateKVSourceData();
                //dataSource.Add("事业1部", 1);
                //dataSource.Add("事业2部", 2);
                //dataSource.Add("事业3部", null);

                args.Model.部门.KVSource = args.Model.部门.CreateKVSource().AddRange(dataSource);
                args.Model.部门2.KVSource = args.Model.部门2.CreateKVSource().AddRange(dataSource);

                var list = EPPlusHelper.GetList(args).ToList();

                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
            [KVSet("'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门 { get; set; }
            [KVSet(false)]
            public KV<string, long?> 部门2 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       Helper.GetEquals_KV(this.部门, y.部门) &&
                       Helper.GetEquals_KV(this.部门2, y.部门2);
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       Helper.GetHashCode_KV(this.部门) +
                       Helper.GetHashCode_KV(this.部门2);
            }
        }
    }
}
