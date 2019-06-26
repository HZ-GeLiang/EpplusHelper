﻿using System;
using System.Data;
using System.IO;
using EPPlusExtensions;
using OfficeOpenXml;

namespace SampleApp.Core
{
    /// <summary>
    /// 获得模版数据检测提示.
    /// </summary>
    class Sample02_6
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_1.xlsx";
            using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                var args = EPPlusHelper.GetExcelListArgsDefault<DataRow>(ws, 2);
                args.WhereFilter = a => Convert.ToInt32(a["序号"]) <= 3;
                args.HavingFilter = a => a["部门负责人"].ToString() == "赵六";
                var dt = EPPlusHelper.GetDataTable(args);
                Console.WriteLine("读取完毕");
            }

            Console.ReadKey();
        }
    }
}
