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
    class Sample02_7
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_7.xlsx";
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                try
                {
                    var args = EPPlusHelper.GetExcelListArgsDefault<userLeaveInfoStat>(ws, 3);
                    var list = EPPlusHelper.GetList<userLeaveInfoStat>(args);
                    ObjectDumper.Write(list);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                Console.WriteLine("读取完毕");
            }

            Console.ReadKey();
        }


        public class userLeaveInfoStat
        {
            public string 序号 { get; set; }
            public string 姓名 { get; set; }
            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public string 请假次数1 { get; set; }
            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public string 请假次数2 { get; set; }
        }

    }
}
