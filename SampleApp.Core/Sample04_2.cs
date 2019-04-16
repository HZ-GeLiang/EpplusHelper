using System.IO;
using EPPlusExtensions;

namespace SampleApp.Core
{
    /// <summary>
    /// 自动初始化填充配置
    /// </summary>
    class Sample04_2
    {
        public void Run()
        {
            string filePath = $@"模版\Sample04_2.xlsx";
            string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));
            var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileOutDirectoryName);
            var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result";
            foreach (var item in defaultConfigList)
            {
                //将字符串全部写入文件
                File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDateTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDateTableSnippe);
                File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
            }

            OpenDirectoryHelp.OpenFilePath(System.IO.Path.Combine(OpenDirectoryHelp.GetSaveFilePath(), @"Debug\模版\"));
        }
    }
}
