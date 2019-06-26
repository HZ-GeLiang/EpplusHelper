using System;
using System.IO;
using EPPlusExtensions;
using OfficeOpenXml;

namespace SampleApp.Core
{
    /// <summary>
    /// 读取数据属性列名时自动重命名
    /// </summary>
    class Sample02_4
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_4.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<Test02_3>(ws, 2);
                args.POCO_Property_AutoRename_WhenRepeat = true;
                args.POCO_Property_AutoRenameFirtName_WhenRepeat = false;
                var list = EPPlusHelper.GetList<Test02_3>(args);
                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class Test02_3
    {

        public string 名字 { get; set; }
        public string 名字2 { get; set; }
        public string 名字3 { get; set; }
    }
}
