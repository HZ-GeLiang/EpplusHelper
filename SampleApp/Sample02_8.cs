using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;
using EPPlusExtensions.Attributes;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_8
    {
        public void Run()
        {
            var errorMsg = EPPlusHelper.GetListErrorMsg(() =>
            {
                string filePath = @"模版\Sample02_7.xlsx";
                using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<Sample02_7.userLeaveInfoStat>(ws, 3);
                    args.GetList_NeedAllException = true;
                    args.GetList_ErrorMessage_OnlyShowColomn = true;
                    var list = EPPlusHelper.GetList<Sample02_7.userLeaveInfoStat>(args);
                    //ObjectDumper.Write(list);

                    Console.WriteLine("读取完毕");
                }
            });

            if (errorMsg?.Length > 0)
            {
                Console.WriteLine(errorMsg);
                Console.ReadKey();
                return;
            }

            Console.ReadKey();
        }

    }
}
