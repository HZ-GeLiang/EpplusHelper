﻿using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Linq;

namespace Core50
{

    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\gl\Desktop\02\637813937556790980.xlsx";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var arg = EPPlusHelper.GetExcelListArgsDefault<Sheet1>(ws, 2);
                arg.ScanLine = ScanLine.SingleLine;
                var list = EPPlusHelper.GetList(arg).ToList();
         
                Console.WriteLine("读取完毕");
                
            }

            Console.WriteLine("Hello World!");
        }
    }
}
