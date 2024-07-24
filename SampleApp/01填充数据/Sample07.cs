using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.IO;

namespace SampleApp._01填充数据
{
    public class Sample07
    {
        public static bool OpenDir = true;
        public static string FilePathSave = @"模版\01填充数据\ResultSample07.xlsx";

        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample07.xlsx";
            //var wsName = 1;
            using (var ms = new MemoryStream())
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var worksheet = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);

                ExcelRange cell = worksheet.Cells["A2"];

                EPPlusHelper.SetWorksheetCellValue(cell, "设置值给富文本单元格");
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(FilePathSave);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }
    }
}