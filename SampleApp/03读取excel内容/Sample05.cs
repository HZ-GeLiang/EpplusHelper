using System;
using System.Collections.Generic;
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
    public class Sample05
    {
        public static IEnumerable<ExcelModel> Run()
        {
            var dataSource = new Dictionary<string, long?>();
            dataSource.Add("事业1部", 1);
            dataSource.Add("事业2部", 2);
            dataSource.Add("事业3部", null);

            return Run(dataSource);
        }

        public static List<ExcelModel> Run(Dictionary<string, long?> dataSource)
        {
            string filePath = @"模版\03读取excel内容\Sample05.xlsx";
            var wsName = 1;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                args.ScanLine = ScanLine.SingleLine;

                var propModel = new ExcelModel();

                //var dataSource = propModel.部门2.CreateKVSourceData();
                //dataSource.Add("事业1部", 1);
                //dataSource.Add("事业2部", 2);
                //dataSource.Add("事业3部", null);

                args.KVSource.Add(nameof(propModel.部门), propModel.部门.CreateKVSource().AddRange(dataSource));
                args.KVSource.Add(nameof(propModel.部门2), propModel.部门2.CreateKVSource().AddRange(dataSource));

                var list = EPPlusHelper.GetList<ExcelModel>(args);

                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
            [KVSet("部门", true, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门 { get; set; }
            [KVSet("部门", false, "'{0}'在数据库中未找到", "部门2")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门2 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门.Key == y.部门.Key &&
                       this.部门.Value == y.部门.Value &&
                       this.部门.HasValue == y.部门.HasValue &&
                       this.部门2.Key == y.部门2.Key &&
                       this.部门2.Value == y.部门2.Value &&
                       this.部门2.HasValue == y.部门2.HasValue;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.Key.GetHashCode() +
                       this.部门.Value.GetHashCode() +
                       this.部门.HasValue.GetHashCode() +
                       this.部门2.Key.GetHashCode() +
                       this.部门2.Value.GetHashCode() +
                       this.部门2.HasValue.GetHashCode();
            }
        }
    }
}
