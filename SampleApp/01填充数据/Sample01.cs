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
    class Sample01
    {
        public void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            string filePathSave = @"模版\01填充数据\ResultSample01.xlsx";
            var wsName = "基本填充";
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                configSource.Body[1].Option.DataSource = Sample00.GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);

                #region 添加密码
                //var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");
                //ws.Protection.IsProtected = true;
                //ws.Protection.AllowSelectLockedCells = false;
                //ws.Protection.AllowSelectUnlockedCells = true;
                //ws.Protection.SetPassword("123");
                #endregion

                //删除的2种方式
                //EPPlusHelper.DeleteWorksheet(excelPackage, wsName);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(filePathSave);
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }
    }
}
