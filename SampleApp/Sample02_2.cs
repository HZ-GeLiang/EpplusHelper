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
using System.ComponentModel.DataAnnotations;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_2
    {
        public void Run()
        {
            string tempPath = @"模版\Sample02_2.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                List<ysbm2> list=  EpplusHelper.GetList<ysbm2>(ws, 2);
                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class ysbm2
    {
        public string 序号 { get; set; }
        [Required(ErrorMessage = "部门不允许为空")]
        public string 部门 { get; set; }

        [Required(ErrorMessage = "部门Id不能为空")]
        [StringLength(5, ErrorMessage = "部门Id长度要在3-5之间", MinimumLength = 3)]
        [Range(101, 99999, ErrorMessage = "值必须在[101,99999]之间")]
        //public string 部门Id { get; set; }
        public long 部门Id { get; set; }

        public string 预算部门 { get; set; }
        public string 预算部门负责人 { get; set; }
        public string 部门负责人 { get; set; }
        public string 部门负责人确认签字 { get; set; }
    }



}