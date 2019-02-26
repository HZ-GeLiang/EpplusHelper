﻿using EpplusExtensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using EpplusExtensions.Attributes; 

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_3
    {
        public void Run()
        {
            string tempPath = @"模版\Sample02_3.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, 1);
                List<PeopleInfo> list = EpplusHelper.GetList<PeopleInfo>(ws, 2);
                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class PeopleInfo
    {
        public string 序号 { get; set; }
        [Unique()]
        public string 名字 { get; set; }
        [EnumUndefined("{0}的性别'{1}'填写不正确", "名字", "性别")]
        public Gender? 性别 { get; set; }
        public DateTime? 出生日期 { get; set; }
        public string 身份证号码 { get; set; }
        public int 年龄 { get; set; }
    }

    public enum Gender
    {
        男 = 1,
        女 = 2,
        未知 = 3,
    }

}