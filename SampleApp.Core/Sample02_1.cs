using System;
using System.Collections.Generic;
using System.IO;
using EPPlusExtensions;
using OfficeOpenXml;

namespace SampleApp.Core
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_1
    {
        public void Run()
        {
            ReadLine(2, ScanLine.MergeLine);
            Console.WriteLine("==========================");
            ReadLine(2, ScanLine.SingleLine);//excel的5,6行是合并的,用SingleLine读取,那么第6行的数据是第5行的
            Console.WriteLine("=========================="); 
            ReadLine(10, ScanLine.SingleLine);

            int a = 3;
        }

        public static void ReadLine(int rowIndex, ScanLine scanLine)
        {

            string tempPath = @"模版\Sample02_1.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                List<ysbm> list;
                try
                {
                    var args = EPPlusHelper.GetExcelListArgsDefault<ysbm>(ws, rowIndex);
                    args.ScanLine = scanLine;

                    if (rowIndex != 2) args.RowIndex_DataName = 1; //这个if 仅针对与当前Demo写的

                    list = EPPlusHelper.GetList<ysbm>(args);
                    ObjectDumper.Write(list);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class ysbm
    {
        public string 序号 { get; set; }
        public string 部门 { get; set; }
        public string 部门负责人 { get; set; }
        public string 部门负责人确认签字 { get; set; }
    }
}
