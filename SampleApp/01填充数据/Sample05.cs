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

namespace SampleApp._01填充数据
{
    public class Sample05
    {
        public static bool OpenDir = true;
        public static string filePathSave = @"模版\01填充数据\ResultSample05.xlsx";
        public static void Run()
        {
            string filePath = @"模版\01填充数据\Sample05.xlsx";
            var wsName = 1;
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var config = EPPlusHelper.GetEmptyConfig();
                var configSource = EPPlusHelper.GetEmptyConfigSource();
                EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, wsName);
                configSource.Head["budgetCycle"] = "上半年";
                configSource.Body[1].Option.DataSource = GetDataTable_Body();
                configSource.Body[1].Option.FillMethod = new SheetBodyFillDataMethod();

                //配置项只到G列,但是H列还有公式,需要自己添加,如果不添加,在处理样式时,H列的公式将会没有
                config.Body[1].Option.ConfigLine.Add(new EPPlusConfigFixedCell { Address = "H3" });//没办法在 SetConfigBodyFromExcel() 的 configLine中添加,需要自己写

                EPPlusHelper.FillData(excelPackage, config, configSource, "预算", 1);
                EPPlusHelper.DeleteWorksheetAll(excelPackage, EPPlusHelper.FillDataWorkSheetNameList);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save(filePathSave);
            }
            if (OpenDir)
            {
                System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
            }
        }
        static DataTable GetDataTable_Body()
        {
            var dtBody = new DataTable();
            dtBody.Columns.Add("Name");
            dtBody.Columns.Add("科目");
            dtBody.Columns.Add("静态预算");
            dtBody.Columns.Add("追加预算");
            dtBody.Columns.Add("已冻结");
            dtBody.Columns.Add("实扣");
            //dtBody.Columns.Add("a");

            var dr = dtBody.NewRow();
            dr["Name"] = "董事办";
            dr["科目"] = "业务费用-礼品费";
            dr["静态预算"] = 12345;
            dr["追加预算"] = 0;
            dr["已冻结"] = 0;
            dr["实扣"] = 345;
            dtBody.Rows.Add(dr);

            dr = dtBody.NewRow();
            dr["Name"] = "董事办";
            dr["科目"] = "业务费用-招待费";
            dr["静态预算"] = 12345;
            dr["追加预算"] = 0;
            dr["已冻结"] = 0;
            dr["实扣"] = 2345;
            dtBody.Rows.Add(dr);

            return dtBody;
        }
    }
}
