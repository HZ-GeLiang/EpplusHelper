﻿using EPPlusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample01
    {
        //逐行读取
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample01.xlsx";
            var wsName = "逐行读取";
            using (var fs = EPPlusHelper.GetFileStream(filePath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);
                args.ScanLine = ScanLine.SingleLine;
                var list = EPPlusHelper.GetList(args).ToList();//excel的5,6行是合并的,用SingleLine读取,list的第4,5条数据内容是一样的
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public int 序号 { get; set; }
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