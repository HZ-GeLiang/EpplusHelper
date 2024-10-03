using EPPlusExtensions;
using OfficeOpenXml;
using System.IO;

namespace SampleApp._01填充数据
{
    public class Sample07
    {
        public static bool OpenDir = true;
        public static string filePathSave = @"模版\01填充数据\ResultSample07.xlsx";

        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample07.xlsx";
            //var wsName = 1;

            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var worksheet = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);

                ExcelRange cell = worksheet.Cells["A2"];

                EPPlusHelper.SetWorksheetCellValue(cell, "设置值给富文本单元格");
                EPPlusHelper.Save(excelPackage, filePathSave);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }
    }
}