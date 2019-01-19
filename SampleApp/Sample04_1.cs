using EpplusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp
{
    /// <summary>
    /// 填充数据与数据源同步
    /// </summary>
    class Sample04_1
    {
        public void Run()
        {

            //string tempPath = $@"模版\各项导出模板.xlsx";
            string tempPath = $@"C:\Users\child\Desktop\11111.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                var dict = EpplusHelper.FillExcelDefaultConfig(excelPackage, new Dictionary<int, int>()
                {
                    //{1,2},
                    //{2,2},
                    //{3,2},
                    //{4,2},
                    //{5,2},
                    //{6,1},
                    {1,1},
                });
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                //ms.Save(@"模版\各项导出模板_Result.xlsx");
                ms.Save($@"C:\Users\child\Desktop\aaa.xlsx");
                foreach (var item in dict)
                {
                    //File.WriteAllText($@"模版\各项导出模板_Result_snippet_{item.Key}.txt", item.Value); //将字符串全部写入文件
                }
            }

            System.Diagnostics.Process.Start(Path.GetDirectoryName(tempPath));
        }

    }
}
