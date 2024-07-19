using EPPlusExtensions;
using OfficeOpenXml;
using SampleApp.MethodExtension;
using System.Data;
using System.IO;

namespace SampleApp._01填充数据
{
    public class Sample06
    {
        public static bool OpenDir = true;
        public static string filePathSave = @"模版\01填充数据\ResultSample06.xlsx";

        public static void Run()
        {
            using (var ms = new MemoryStream())
            using (var excelPackage = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("学生信息");

                var dt = GetDataTable();

                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);

                configSource.Body[1].Option.DataSource = dt;
                configSource.Body[1].Option.FillMethod = new SheetBodyFillDataMethod()
                {
                    FillDataMethodOption = SheetBodyFillDataMethodOption.SynchronizationDataSource,
                    SynchronizationDataSource = new SynchronizationDataSourceConfig()
                    {
                        NeedBody = true,
                        NeedTitle = true,
                        Exclude = "Id"
                    }
                };
                config.Body[1].Option.CustomSetValue = (customValue) =>
                {
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
                        customValue.Cell.Value = config.UseFundamentals
                            ? config.CellFormatDefault(customValue.ColName, customValue.Value, customValue.Cell)
                            : customValue.Value;
                    }
                };

                EPPlusHelper.FillData(config, configSource, worksheet);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(filePathSave);

                // //Add the headers
                // worksheet.Cells[1, 1].Value = "ID";

                // // set some document properties
                //excelPackage.Workbook.Properties.Title = "1";
                //excelPackage.Workbook.Properties.Author = "2";
                // excelPackage.Workbook.Properties.Comments = "33";

                // // set some extended property values
                // excelPackage.Workbook.Properties.Company = "44.";

                // // set some custom property values
                // excelPackage.Workbook.Properties.SetCustomPropertyValue("55", "66");

                // //@"模版\01填充数据\ResultSample05.xlsx";
                // var xlFile = new FileInfo("D:/sample1.xlsx");
                // // save our new workbook in the output directory and we are done!
                // excelPackage.SaveAs(xlFile);

                //var ccc = xlFile.FullName;
            }
        }

        private static DataTable GetDataTable()
        {
            var dtBody = new DataTable();
            dtBody.Columns.Add("Id");
            dtBody.Columns.Add("Name");
            dtBody.Columns.Add("Sex");
            dtBody.Columns.Add("Age");
            dtBody.Columns.Add("Height");

            var dr = dtBody.NewRow();
            dr["Id"] = 1;
            dr["Name"] = "bob";
            dr["Sex"] = "男";
            dr["Age"] = 20;
            dr["Height"] = 170;
            dtBody.Rows.Add(dr);

            dr = dtBody.NewRow();
            dr["Id"] = 2;
            dr["Name"] = "alice";
            dr["Sex"] = "女";
            dr["Age"] = 16;
            dr["Height"] = 166;
            dtBody.Rows.Add(dr);

            return dtBody;
        }
    }
}
