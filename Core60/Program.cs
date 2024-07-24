using EPPlusExtensions;
using OfficeOpenXml;

namespace Core60
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string filePath = @"D:\新建 Microsoft Excel 工作表.xlsx";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var arg = EPPlusHelper.GetExcelListArgsDefault<Sheet1>(ws, 2);
                arg.ScanLine = ScanLine.SingleLine;
                var list = EPPlusHelper.GetList(arg).ToList();

                Console.WriteLine("读取完毕");
            }

            Console.WriteLine("Hello World!");
        }
    }
}