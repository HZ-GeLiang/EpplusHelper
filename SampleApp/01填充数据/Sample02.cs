using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;

namespace SampleApp._01填充数据
{
    class Sample02
    {
        public void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            string filePathSave = @"模版\01填充数据\ResultSample02.xlsx";
            var wsName = "带有标题行的填充";
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                configSource.Head = Sample00.GetDataTable_Head();
                configSource.Body[1].Option.DataSource = Sample00.GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);
                
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(filePathSave);
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }
    }
}
