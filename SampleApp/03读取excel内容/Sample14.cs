using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp._03读取excel内容
{
    public class Sample14
    {
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample14.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                //args.ScanLine = ScanLine.MergeLine;//默认的
                args.ScanLine = ScanLine.SingleLine;
                var list = EPPlusHelper.GetList(args);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }


        public class ExcelModel
        {
            public string A { get; set; }
            public string B { get; set; }
            public string C { get; set; }
            public string D { get; set; }
            public string E { get; set; }
            public string F { get; set; }
            public string G { get; set; }
            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.A == y.A &&
                       this.B == y.B &&
                       this.C == y.C &&
                       this.D == y.D &&
                       this.E == y.E &&
                       this.F == y.F &&
                       this.G == y.G;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.A.GetHashCode() +
                       this.B.GetHashCode() +
                       this.C.GetHashCode() +
                       this.D.GetHashCode() +
                       this.E.GetHashCode() +
                       this.F.GetHashCode() +
                       this.G.GetHashCode();
            }
        }
    }
}
