﻿using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.Data;
using System.IO;

namespace SampleApp._01填充数据
{
    public class Sample04
    {
        public static bool OpenDir = true;
        public static string FilePathSave = @"模版\01填充数据\ResultSample04.xlsx";

        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            var wsName = "带标题行且填充列是单行多列的合并单元格";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                config.Body[1].Option.InsertRowStyle.Operation = InsertRowStyleOperation.CopyStyleAndMergeCell;//表格框线的显示
                config.Body[1].Option.InsertRowStyle.NeedMergeCell = true;//单元格的合并
                configSource.Head = GetDataTable_Head();
                configSource.Body[1].Option.DataSource = GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);

                EPPlusHelper.DeleteWorkSheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);

                using (var ms = EPPlusHelper.GetMemoryStream(excelPackage))
                {
                    ms.Save(FilePathSave);
                }
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