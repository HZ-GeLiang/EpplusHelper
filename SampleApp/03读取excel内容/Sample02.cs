﻿using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample02
    {
        //合并行读取
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample02.xlsx";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                //args.ScanLine = ScanLine.MergeLine;//默认的
                var list = EPPlusHelper.GetList(args).ToList();//输出的是看到的
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public string 序号 { get; set; }
            public string 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj is null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       this.部门 == y.部门 &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       this.部门.GetHashCode() +
                       this.部门负责人.GetHashCode() +
                       this.部门负责人确认签字.GetHashCode();
            }
        }
    }
}