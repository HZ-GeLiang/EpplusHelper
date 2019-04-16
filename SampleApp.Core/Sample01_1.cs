using System.Data;
using System.IO;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.Core.MethodExtension;

namespace SampleApp.Core
{
    /// <summary>
    /// 基本用法
    /// </summary>
    class Sample01_1
    {
        public void Run()
        {
            string tempPath = @"模版\Sample01_1.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, 1);
                var dtHead = GetDataTable_Head();
                EPPlusHelper.SetConfigSourceHead(configSource, dtHead, dtHead.Rows[0]);
                configSource.SheetBody[1] = GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", 1);
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");
                ws.Protection.IsProtected = true;
                ws.Protection.AllowSelectLockedCells = false;
                ws.Protection.AllowSelectUnlockedCells = true;
                ws.Protection.SetPassword("123");
                EPPlusHelper.DeleteWorksheet(excelPackage, 1);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample01_1_Result.xlsx");
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(tempPath));
        }

        static DataTable GetDataTable_Head()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Title");

            DataRow dr = dt.NewRow();
            dr["Title"] = "2018第一学期考试";
            dt.Rows.Add(dr);
            return dt;

        }
        static DataTable GetDataTable_Body()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Chinese");
            dt.Columns.Add("Math");
            dt.Columns.Add("English");

            DataRow dr = dt.NewRow();
            dr["Name"] = "张三";
            dr["Chinese"] = 60;
            dr["Math"] = 60.5;
            dr["English"] = 61;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Name"] = "李四";
            dr["Chinese"] = 70;
            dr["Math"] = 80.5;
            dr["English"] = 91;
            dt.Rows.Add(dr);

            return dt;

        }
    }
}
