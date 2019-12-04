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
    class Sample04
    {
        public void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            string filePathSave = @"模版\01填充数据\ResultSample04.xlsx";
            var wsName = "带标题行且填充列是单行多列的合并单元格";
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                config.Body[1].Option.InsertRowStyle.Operation = InsertRowStyleOperation.CopyStyleAndMergeCell;//表格框线的显示
                config.Body[1].Option.InsertRowStyle.NeedMergeCell = true;//单元格的合并
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
