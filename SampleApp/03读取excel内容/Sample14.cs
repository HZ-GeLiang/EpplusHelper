using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample14
    {
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample14.xlsx";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                //args.ScanLine = ScanLine.MergeLine;//默认的
                var list = EPPlusHelper.GetList(args).ToList();
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }


        public class ExcelModel
        {
            public int A { get; set; }
            public int B { get; set; }
            public int C { get; set; }
            public int D { get; set; }
            public int E { get; set; }
            public int F { get; set; }
            public int G { get; set; }
            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
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
