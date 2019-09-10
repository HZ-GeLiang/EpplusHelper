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
    class Sample01_1
    {
        public void Run()
        {
            string filePath = @"模版\Sample01_1.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath)) 
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                //EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, 1);
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, "Sheet2");
                config.Body[1].Option.InsertRowStyle.Operation = InsertRowStyleOperation.CopyStyleAndMergeCell;
                config.Body[1].Option.InsertRowStyle.NeedMergeCell = false;
                var dtHead = GetDataTable_Head();
                //EPPlusHelper.SetConfigSourceHead(configSource, dtHead, dtHead.Rows[0]);
                //EPPlusHelper.SetConfigSourceHead(configSource, dtHead);
                configSource.Head = dtHead;


                //configSource.Head["budgetCycle"] = "上半年";

                //configSource.Body.ConfigList = new List<EPPlusConfigSourceBodyConfig>()
                //{
                //    new EPPlusConfigSourceBodyConfig
                //    {
                //        Nth = 1,
                //        Option = new EPPlusConfigSourceBodyOption()
                //        {
                //            DataSource = GetDataTable_Body()
                //        }
                //    }
                //};

                configSource.Body[1].Option.DataSource = GetDataTable_Body();


                var stopwatch = new System.Diagnostics.Stopwatch();
                Console.WriteLine("runTime 开始");
                stopwatch.Start();

                //EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", 1);
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", "Sheet2");
                stopwatch.Stop();
                Console.WriteLine("导出数据runTime 时差:" + stopwatch.Elapsed);
                Console.WriteLine("导出数据runTime 毫秒:" + stopwatch.ElapsedMilliseconds);

                stopwatch.Reset();
                stopwatch.Start();
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");
                //ws.Protection.IsProtected = true;
                //ws.Protection.AllowSelectLockedCells = false;
                //ws.Protection.AllowSelectUnlockedCells = true;
                //ws.Protection.SetPassword("123");
                //EPPlusHelper.DeleteWorksheet(excelPackage, 1);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNames);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample01_1_Result.xlsx");
                stopwatch.Stop();
                Console.WriteLine("保存数据runTime 时差:" + stopwatch.Elapsed);
                Console.WriteLine("保存数据runTime 毫秒:" + stopwatch.ElapsedMilliseconds);

                Console.ReadKey();
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
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

            //for (int i = 0; i < 400*c*c; i++)
            //for (int i = 0; i < 1000000; i++)
            //for (int i = 0; i < 800000; i++)
            //for (int i = 0; i < 100000; i++)
            //for (int i = 0; i < 1048576/2-1; i++)
            //for (int i = 0; i < 550000; i++)
            for (int i = 0; i < 550; i++)
            {
                DataRow dr = dt.NewRow();
                dr["Name"] = $"张三{i + 1}";
                dr["Chinese"] = 60;
                dr["Math"] = 60.5;
                dr["English"] = 61;
                dt.Rows.Add(dr);
            }

            //DataRow dr = dt.NewRow();
            //dr["Name"] = "张三";
            //dr["Chinese"] = 60;
            //dr["Math"] = 60.5;
            //dr["English"] = 61;
            //dt.Rows.Add(dr);

            //dr = dt.NewRow();
            //dr["Name"] = "李四";
            //dr["Chinese"] = 70;
            //dr["Math"] = 80.5;
            //dr["English"] = 91;
            //dt.Rows.Add(dr);

            return dt;

        }
    }
}
