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

namespace SampleApp._03读取excel内容
{
    class Sample12
    {
        public void Run()
        {
            var errorMsg = EPPlusHelper.GetListErrorMsg(() =>
            {
                string filePath = @"模版\03读取excel内容\Sample12.xlsx";
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<UserLeaveStat>(ws, 3);
                    args.GetList_NeedAllException = true; //默认false
                    args.GetList_ErrorMessage_OnlyShowColomn = true; //默认false
                                                                     //true:  参数名: 姓名(B列)
                                                                     //false: 参数名: 姓名(B3,B4,B5)

                    var list = EPPlusHelper.GetList<UserLeaveStat>(args);
                    ObjectDumper.Write(list);
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
        class UserLeaveStat
        {
            public string 序号 { get; set; }
            public string 姓名 { get; set; }
            //public int 姓名 { get; set; } //要效果,取消注释

            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public string JanuaryStatistics { get; set; }

            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public string FebruaryStatistics { get; set; }
        }
    }
}
