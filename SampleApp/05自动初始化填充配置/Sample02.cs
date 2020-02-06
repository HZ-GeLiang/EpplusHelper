using EPPlusExtensions;
using System.IO;

namespace SampleApp._05自动初始化填充配置
{
    public class Sample02
    {
        public static bool OpenDir = true;
        public static void Run()
        {
            string filePath = @"模版\05自动初始化填充配置\Sample02.xlsx";

            string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));
            var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileOutDirectoryName, null);
            var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result";
            foreach (var item in defaultConfigList)
            {
                //将字符串全部写入文件
                File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDataTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDataTableSnippe);
                File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
                // OpenDirectoryHelp.OpenFilePath(System.IO.Path.Combine(OpenDirectoryHelp.GetSaveFilePath(), @"Debug\模版\05自动初始化填充配置\"));
            }
        }
    }
}
