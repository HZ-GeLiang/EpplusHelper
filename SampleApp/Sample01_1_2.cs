using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;

namespace SampleApp
{
    /// <summary>
    /// 基本用法
    /// </summary>
    class Sample01_1_2
    {
        public void Run()
        {
            string filePath = @"模版\Sample01_1_2.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, 1);
                //EPPlusHelper.SetConfigSourceHead(configSource, dtHead, dtHead.Rows[0]);
                //EPPlusHelper.SetConfigSourceHead(configSource, dtHead);
                configSource.Head = Sample01_1.GetDataTable_Head();
                configSource.Body[1].Option.DataSource = Sample01_1.GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", 1);


                //删除的2种方式
                //EPPlusHelper.DeleteWorksheet(excelPackage, 1);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample01_1_2_Result.xlsx");
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }


    }
}
