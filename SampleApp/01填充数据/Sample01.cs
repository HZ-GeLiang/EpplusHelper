using System;
using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.Data;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Style;

namespace SampleApp._01填充数据
{
    public class Sample01
    {
        public static bool OpenDir = true;
        public static string FilePathSave = @"模版\01填充数据\ResultSample01.xlsx";

        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample01.xlsx";
            var wsName = "基本填充";
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                configSource.Body[1].Option.DataSource = GetDataTable_Body();
                EPPlusHelper.FillData(excelPackage, config, configSource, "导出测试", wsName);

                #region 添加密码
                //var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");
                //ws.Protection.IsProtected = true;
                //ws.Protection.AllowSelectLockedCells = false;
                //ws.Protection.AllowSelectUnlockedCells = true;
                //ws.Protection.SetPassword("123");
                #endregion

                //删除的2种方式
                //EPPlusHelper.DeleteWorksheet(excelPackage, wsName);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);

                var worksheet = EPPlusHelper.GetExcelWorksheet(excelPackage, "导出测试");

                //自动列宽
                //worksheet.Cells["D1"].AutoFitColumns();
                //worksheet.Column(4).AutoFit(0); 

               worksheet.Cells.Style.ShrinkToFit = true;//单元格自动适应大小  (效果是:单元格大小不变, 缩放里面的文字)


                worksheet.Cells["D1"].Merge = true;
                worksheet.Cells["D1"].Style.WrapText = true;
                //  worksheet.Row(1).Height = 355;//设置行高
                worksheet.Row(1).CustomHeight = true;//自动调整行高
                worksheet.Column(4).BestFit = true;//当Bestfit设置为true时，当用户在单元格中输入数字时，该列将变宽

                
                var cell = worksheet.Cells["D1"];
                worksheet.Row(1).Height = MeasureTextHeight(cell.Value.ToString(), cell.Style.Font, 100);

                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(FilePathSave);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }
        public static double MeasureTextHeight(string text, ExcelFont font, int width)
        {
            if (string.IsNullOrEmpty(text)) return 0.0;
            var bitmap = new Bitmap(1, 1);
            var graphics = Graphics.FromImage(bitmap);

            var pixelWidth = Convert.ToInt32(width * 7.5); //7.5 pixels per excel column width 
            var drawingFont = new Font(font.Name, font.Size);
            var size = graphics.MeasureString(text, drawingFont, pixelWidth);

            //72 DPI and 96 points per inch. Excel height in points with max of 409 per Excel requirements. 
            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96, 409);
        }

        static DataTable GetDataTable_Body()
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
