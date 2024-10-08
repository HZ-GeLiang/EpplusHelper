﻿using EPPlusExtensions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Data;
using System.Drawing;
using System.IO;

namespace SampleApp._02填充图片
{
    // 我也没捣鼓出来怎么添加合适
    public class Sample01
    {
        private const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;

        public static void Run()
        {
            string filePath = @"模版\02填充图片\Sample01.xlsx";
            string filePathSave = @"模版\02填充图片\ResultSample01.xlsx";
            var wsName = 1;
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var dtHead = GetDataTable_Head();
                configSource.Head = dtHead;
                configSource.Body[1].Option.DataSource = GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);
                EPPlusHelper.DeleteWorksheet(excelPackage, 1);

                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");

                #region 测试,没写好

                MemoryStream img_ms = CaptchaGen.ImageFactory.GenerateImage("你好中国", 80, 100, 13, 5);//这个是图片的px
                byte[] stream = img_ms.ToArray();
                var bitmap = new Bitmap(img_ms);
                ExcelPicture pic = ws.Drawings.AddPicture("pic1", bitmap);

                //1px = 0.0265cm
                pic.SetPosition(19 * 3, 0);//x ,y的px坐标

                var aaa = ws.Row(4).Height;
                //Console.ReadKey();
                #endregion

                //0.08 = 1px  >1px = 0.08*N -0.01
                ws.Column(3).Width = 0.125 * 10;//excel的单位
                ws.Column(4).Width = 0.125 * 1;//excel的单位
                                               //ws.Column(2).Width =62.1 ;//excel的单位
                                               //ws.Column(3).Width =62.2 ;//excel的单位.3
                                               //ws.Column(4).Width =62.3 ;//excel的单位61.63
                                               //ws.Column(5).Width =62.4 ;//excel的单位61.75
                                               //ws.Column(6).Width =62.5 ;//excel的单位
                                               //17 的16.5(137px)  17.75 的17.13(142px)
                                               //0.126*100 =       12=101px
                                               //0.126*500 =       62.38=504px
                                               //61.75 = 499

                //excel的单位, 如果要从px 转换,那么就 * 0.75 . 注:
                // 建议使用 像素 * 0.75,  内部怎么转换的不清楚. 试了下, 写8.23 得 8.25 (11px ), 譬如 50 得 49.5 (66px)
                ws.Row(4).Height = 0.75 * 80;
                ws.Row(4).Height = 11 * 0.75;
                //ws.Row(5).Height = 50;
                //ws.Cells["A4:A5"].AutoFitColumns(1);

                EPPlusHelper.Save(excelPackage, filePathSave);
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
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
            dt.Columns.Add("Evaluate");

            DataRow dr = dt.NewRow();
            dr["Name"] = "张三";
            dr["Chinese"] = 60;
            dr["Math"] = 60.5;
            dr["English"] = 61;
            dr["Evaluate"] = CaptchaGen.ImageFactory.GenerateImage("合", 50, 100, 13, 0);
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Name"] = "李四";
            dr["Chinese"] = 70;
            dr["Math"] = 80.5;
            dr["English"] = 91;
            dr["Evaluate"] = CaptchaGen.ImageFactory.GenerateImage("优", 50, 100, 13, 0);
            dt.Rows.Add(dr);

            return dt;
        }
    }
}