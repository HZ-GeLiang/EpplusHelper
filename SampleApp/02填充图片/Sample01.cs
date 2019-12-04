using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SampleApp.MethodExtension;

namespace SampleApp._02填充图片
{
    class Sample01
    {
        const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;
        public void Run()
        {
            string filePath = @"模版\02填充图片\Sample01.xlsx";
            string filePathSave = @"模版\02填充图片\ResultSample01.xlsx";
            var wsName = 1;
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var dtHead = Sample00.GetDataTable_Head();
                configSource.Head = dtHead;
                configSource.Body[1].Option.DataSource = Sample00.GetDataTable_Body();
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

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(filePathSave);
            }
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
        }
     
    }
}
