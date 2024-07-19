using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.Data;
using System.IO;

namespace SampleApp._01填充数据
{
    public class Sample02
    {
        public static bool OpenDir = true;
        public static string FilePathSave = @"模版\01填充数据\ResultSample02.xlsx";

        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            var wsName = "带有标题行的填充";
            using (var ms = new MemoryStream())
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                configSource.Head = GetDataTable_Head();
                configSource.Body[1].Option.DataSource = GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);

                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(FilePathSave);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }

        private static DataTable GetDataTable_Head()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Title");
            DataRow dr = dt.NewRow();
            dr["Title"] = "2018第一学期考试";
            dt.Rows.Add(dr);
            return dt;
        }

        private static DataTable GetDataTable_Body()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Chinese");
            dt.Columns.Add("Math");
            dt.Columns.Add("English");
            for (int i = 0; i < 5; i++)
            {
                DataRow dr = dt.NewRow();
                dr["Name"] = $"张三{i + 1}";
                dr["Chinese"] = 60;
                dr["Math"] = 60.5;
                dr["English"] = 61;
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
