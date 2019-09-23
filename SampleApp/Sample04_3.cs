using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using SampleApp.MethodExtension;

namespace SampleApp
{
    /// <summary>
    /// 自动初始化填充配置
    /// </summary>
    class Sample04_3
    {
        public void Run()
        {

            string filePath = $@"模版\Sample04_3.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
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
                ms.Save(@"模版\Sample04_3_Result.xlsx");
                var filePathPrefix = $@"模版\Sample04_3_Result";
                foreach (var item in defaultConfigList)
                {
                    //将字符串全部写入文件
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDataTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDataTableSnippe);
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                }
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));

        }

    }
}
