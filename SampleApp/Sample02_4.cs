using EpplusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp
{
    /// <summary>
    /// 读取数据属性列名时自动重命名
    /// </summary>
    class Sample02_4
    {
        public void Run()
        {
            string tempPath = @"模版\Sample02_4.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = System.IO.File.OpenRead(tempPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage,1);
                var list = EpplusHelper.GetList<Test02_3>(new GetExcelListArgs<Test02_3>()
                {
                    ws = ws,
                    rowIndex_Data = 2,
                    EveryCellPrefix = "",
                    EveryCellReplaceList = null,
                    RowIndex_DataName = 2 - 1,
                    UseEveryCellReplace = true,
                    HavingFilter = null,
                    WhereFilter = null,
                    ReadCellValueOption = ReadCellValueOption.Trim,
                    POCO_Property_AutoRename_WhenRepeat = true,
                    POCO_Property_AutoRenameFirtName_WhenRepeat = false,
                });


                Console.WriteLine("读取完毕");
            }
        }
    }

    internal class Test02_3
    {

        public string 名字 { get; set; }
        public string 名字2 { get; set; }
        public string 名字3 { get; set; }
    }
} 
