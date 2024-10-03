using EPPlusExtensions;
using OfficeOpenXml;
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
        public static string filePathSave = @"模版\05自动初始化填充配置\Sample03_Result.xlsx";

        public static void Run()
        {
            string filePath = @"模版\05自动初始化填充配置\Sample03.xlsx";

            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var dataConfigInfo = new List<ExcelDataConfigInfo>()
                {
                    new ExcelDataConfigInfo() {WorkSheetIndex = 1, TitleLine = 2, TitleColumn = 1},
                    new ExcelDataConfigInfo() {WorkSheetIndex = 2, TitleLine = 2, TitleColumn = 1},
                    new ExcelDataConfigInfo() {WorkSheetIndex = 3, TitleLine = 1, TitleColumn = 1},
                };

                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(excelPackage, dataConfigInfo);
                EPPlusHelper.Save(excelPackage, filePathSave);

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