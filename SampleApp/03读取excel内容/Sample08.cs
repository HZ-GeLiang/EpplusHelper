using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml; 

namespace SampleApp._03读取excel内容
{
    public class Sample08
    {
        public static IEnumerable<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample08.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                args.POCO_Property_AutoRename_WhenRepeat = true;
                args.POCO_Property_AutoRenameFirtName_WhenRepeat = false;
                var list = EPPlusHelper.GetList(args);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public string 名字 { get; set; }
            public string 名字2 { get; set; }
            public string 名字3 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.名字 == y.名字 &&
                       this.名字2 == y.名字2 &&
                       this.名字3 == y.名字3;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.名字.GetHashCode() +
                       this.名字2.GetHashCode() +
                       this.名字3.GetHashCode();
            }
        }
    }
}
