﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
    class Sample03_1_2
    {
        public void Run()
        {
            string filePath = @"模版\Sample03_1.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, "Sheet2");

                configSource.Body[1].Option.DataSource = GetProduct1();
                configSource.Body[1].Option.FillMethod = new SheetBodyFillDataMethod()
                {
                    FillDataMethodOption = SheetBodyFillDataMethodOption.SynchronizationDataSource,
                    SynchronizationDataSource = new SynchronizationDataSourceConfig()
                    {
                        //NeedBody = true,
                        //NeedTitle = true,
                        NeedBody = false,
                        NeedTitle = false,
                        Include = "使用人,购买时间"
                    }
                };
                config.Body[1].Option.CustomSetValue = (customValue) =>
                {
                    //config.Body[1].Option.ConfigLine
                    if (customValue.Area == FillArea.TitleExt)
                    {
                        customValue.Cell.Value = $"标题扩展-{customValue.Value}";
                    }
                    else if (customValue.Area == FillArea.ContentExt)
                    {
                        customValue.Cell.Value = $"内容扩展-{customValue.Value}";

                        customValue.Cell.StyleID = customValue.Worksheet.Cells[4, 4].StyleID;
                    }
                    else
                    {
                        //cell.Value = val;
                        customValue.Cell.Value = config.UseFundamentals ?
                            config.CellFormatDefault(customValue.ColName, customValue.Value, customValue.Cell) 
                            : customValue.Value;
                    }
                };

                configSource.Body[2].Option.DataSource = GetProduct2();
                configSource.Body[3].Option.DataSource = GetProduct3();
                configSource.Body[3].Option.FillMethod = new SheetBodyFillDataMethod()
                {
                    FillDataMethodOption = SheetBodyFillDataMethodOption.SynchronizationDataSource,
                    SynchronizationDataSource = new SynchronizationDataSourceConfig()
                    {
                        //NeedBody = true,
                        //NeedTitle = true,
                        NeedBody = false,
                        NeedTitle = false,
                        Exclude = "Id"
                    }
                };

                //发现样式有问题,用这个设置,但是生成的文件会增加体积
                config.Body[3].Option.InsertRowStyle.Operation = InsertRowStyleOperation.CopyStyleAndMergeCell;

                EPPlusHelper.FillData(excelPackage, config, configSource, "Result", "Sheet2");
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample03_1_2_Result.xlsx");
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
            dr["Wide"] = "11cm";
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

            dr = dt.NewRow();
            dr["Id"] = "8";
            dr["Name"] = "月饼";
            dr["Price"] = "40";
            dr["Weight"] = "5cm";
            dr["Long"] = "6cm";
            dr["Wide"] = "4cm";
            dr["高"] = "3cm";
            dr["经销商"] = "yb";
            dt.Rows.Add(dr);
            return dt;
        }

    }
}
