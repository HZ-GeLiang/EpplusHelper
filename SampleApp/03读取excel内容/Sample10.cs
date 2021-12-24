using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System;
using System.Data;
using System.IO;

namespace SampleApp._03读取excel内容
{
    public class Sample10
    {
        public static DataTable Run()
        {
            string filePath = @"模版\03读取excel内容\Sample01.xlsx";
            var wsName = "逐行读取";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<DataRow>(ws, 2);
                args.WhereFilter = a => Convert.ToInt32(a["序号"]) <= 3;
                args.HavingFilter = a => a["部门负责人"].ToString() == "赵六";
                var dt = EPPlusHelper.GetDataTable(args);
                var txt = dt.ToText();
                Console.WriteLine(txt);
                Console.WriteLine("读取完毕");
                return dt;
            }
        }
    }
}
