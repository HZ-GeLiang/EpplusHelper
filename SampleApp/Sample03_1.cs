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
    /// 填充数据与数据源同步
    /// </summary>
    class Sample03_1
    {
        public void Run()
        {
            string filePath = @"模版\Sample03_1.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, "Sheet1");
                configSource.SheetBody[1] = GetProduct1();
                configSource.SheetBody[2] = GetProduct2();
                configSource.SheetBody[3] = GetProduct3();
                configSource.SheetBodyFillModel.Add(1, new SheetBodyFillDataMethod()
                {
                    FillDataMethodOption = SheetBodyFillDataMethodOption.SynchronizationDataSource,
                    SynchronizationDataSource = new SynchronizationDataSourceConfig()
                    {
                        NeedBody = true,
                        NeedTitle = true,
                        Include = "使用人,购买时间"
                    }
                });
                configSource.SheetBodyFillModel.Add(3, new SheetBodyFillDataMethod()
                {
                    FillDataMethodOption = SheetBodyFillDataMethodOption.SynchronizationDataSource,
                    SynchronizationDataSource = new SynchronizationDataSourceConfig()
                    {
                        NeedBody = true,
                        NeedTitle = true,
                        Exclude = "Id"
                    }
                });
                EPPlusHelper.FillData(excelPackage, config, configSource, "Result", "Sheet1");
                EPPlusHelper.DeleteWorksheet(excelPackage, "Sheet1");
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample03_1_result.xlsx");
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }


        static DataTable GetProduct1()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Price");
            dt.Columns.Add("Qty");
            dt.Columns.Add("使用人");
            dt.Columns.Add("购买时间");
            DataRow dr = dt.NewRow();
            dr["Id"] = "3";
            dr["Name"] = "ThinkPad P1";
            dr["Price"] = "$1406.33";
            dr["Qty"] = "2";
            dr["使用人"] = "张三";
            dr["购买时间"] = "2018-1-1";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Id"] = "4";
            dr["Name"] = "iphone XS";
            dr["Price"] = "￥9999";
            dr["Qty"] = "7";
            dr["使用人"] = "小七";
            dr["购买时间"] = "2018-1-2";
            dt.Rows.Add(dr);
            return dt;
        }
        static DataTable GetProduct2()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Price");
            dt.Columns.Add("Color");
            DataRow dr = dt.NewRow();
            dr["Id"] = "8";
            dr["Name"] = "杯子";
            dr["Price"] = "55";
            dr["Color"] = "红";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Id"] = "9";
            dr["Name"] = "耳机";
            dr["Price"] = "65";
            dr["Color"] = "蓝";
            dt.Rows.Add(dr);
            return dt;
        }
        static DataTable GetProduct3()
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Price");
            dt.Columns.Add("Weight");
            dt.Columns.Add("Long");
            dt.Columns.Add("Wide");
            dt.Columns.Add("高");
            dt.Columns.Add("经销商");
            DataRow dr = dt.NewRow();
            dr["Id"] = "3";
            dr["Name"] = "杯子";
            dr["Price"] = "55";
            dr["Weight"] = "1kg";
            dr["Long"] = "10cm";
            dr["Wide"] = "10cm";
            dr["高"] = "22cm";
            dr["经销商"] = "A";
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Id"] = "8";
            dr["Name"] = "耳机";
            dr["Price"] = "65";
            dr["Weight"] = "1lb";
            dr["Long"] = "13cm";
            dr["Wide"] = "20cm";
            dr["高"] = "30cm";
            dr["经销商"] = "B";
            dt.Rows.Add(dr);
            return dt;
        }

    }
}
