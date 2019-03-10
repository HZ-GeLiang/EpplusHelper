using EpplusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp
{
    /// <summary>
    /// 填充数据与数据源同步
    /// </summary>
    class Sample04_3
    {
        public void Run()
        {

            string tempPath = $@"模版\Sample04_3.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var sheetTitleLineNumber = new Dictionary<int, int>()
                {
                    {1, 2},
                    {2, 2},
                    {3, 1},
                };
                var defaultConfigList = EpplusHelper.FillExcelDefaultConfig(excelPackage, sheetTitleLineNumber);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample04_3_Result.xlsx");
                var filePathPrefix = $@"模版\Sample04_3_Result";
                foreach (var item in defaultConfigList)
                {
                    //将字符串全部写入文件
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDateTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDateTableSnippe);
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                }
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(tempPath));
        }

    }
}
