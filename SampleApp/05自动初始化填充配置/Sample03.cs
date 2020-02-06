using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.Collections.Generic;
using System.IO;

namespace SampleApp._05自动初始化填充配置
{
    /// <summary>
    /// 自动初始化填充配置
    /// </summary>
    public class Sample03
    {
        public static bool OpenDir = true;
        public static string FilePathSave = @"模版\05自动初始化填充配置\ResultSample03.xlsx";
        public static void Run()
        {
            string filePath = @"模版\05自动初始化填充配置\Sample03.xlsx";

            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var dataConfigInfo = new List<ExcelDataConfigInfo>()
                {
                    new ExcelDataConfigInfo() {WorkSheetIndex = 1, TitleLine = 2, TitleColumn = 1},
                    new ExcelDataConfigInfo() {WorkSheetIndex = 2, TitleLine = 2, TitleColumn = 1},
                    new ExcelDataConfigInfo() {WorkSheetIndex = 3, TitleLine = 1, TitleColumn = 1},
                };

                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(excelPackage, dataConfigInfo);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(FilePathSave);
                var filePathPrefix = Path.GetDirectoryName(filePath);
                foreach (var item in defaultConfigList)
                {
                    //将字符串全部写入文件
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDataTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDataTableSnippe);
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                }
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }
    }
}
