using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample15
    {
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample15.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
            public int Num1 { get; set; }
            public int Num2 { get; set; }
            public int Sum { get; set; }
            public int CopyNum1 { get; set; }
            public int CopySum { get; set; }
             
            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.Num1 == y.Num1 &&
                       this.Num2 == y.Num2 &&
                       this.Sum == y.Sum &&
                       this.CopyNum1 == y.CopyNum1 &&
                       this.CopySum == y.CopySum;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.Num1.GetHashCode() +
                       this.Num2.GetHashCode() +
                       this.Sum.GetHashCode() +
                       this.CopyNum1.GetHashCode() +
                       this.CopySum.GetHashCode();
            }
        }
    }
}
