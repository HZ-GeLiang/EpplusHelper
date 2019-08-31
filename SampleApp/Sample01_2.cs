using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System.Windows.Media.Imaging;
using System.Runtime.InteropServices;
using EPPlusExtensions;
using SampleApp.MethodExtension;

namespace SampleApp
{
    /// <summary>
    /// 基本用法,
    /// </summary>
    class Sample01_2
    {

        public void Run()
        {
            string filePath = @"模版\Sample01_2.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, 1);
                var dtHead = GetDataTable_Head();
                EPPlusHelper.SetConfigSourceHead(configSource, dtHead, dtHead.Rows[0]);
                configSource.Body.InfoList = new List<EPPlusConfigSourceBodyInfo>()
                {
                    new EPPlusConfigSourceBodyInfo
                    {
                        Nth = 1,
                        Option = new EPPlusConfigSourceBodyOption()
                        {
                            DataSource = GetDataTable_Body()
                        }
                    }
                };
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", 1);
                EPPlusHelper.DeleteWorksheet(excelPackage, 1);

                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");

                MemoryStream img_ms = CaptchaGen.ImageFactory.GenerateImage("你好中国", 80, 100, 13, 5);//这个是图片的px
                byte[] stream = img_ms.ToArray();
                var bitmap = new Bitmap(img_ms);
                ExcelPicture pic = ws.Drawings.AddPicture("pic1", bitmap);

                //1px = 0.0265cm
                pic.SetPosition(19 * 3, 0);//x ,y的px坐标

                var aaa = ws.Row(4).Height;

                //Console.ReadKey();

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

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(@"模版\Sample01_2_result.xlsx");
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }
        const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;

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
