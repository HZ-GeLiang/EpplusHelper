using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_1_3
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_1.xlsx";
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                try
                {
                    ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet2");
                    var args = EPPlusHelper.GetExcelListArgsDefault<ysbm>(ws, 2);
                    args.ScanLine = ScanLine.SingleLine;

                    var propModel = new ysbm();

                    var dataSource = propModel.部门2.CreateKVSourceData();
                    dataSource.Add("事业1部", 1);
                    dataSource.Add("事业2部", 2);
                    dataSource.Add("事业3部", null);
                    args.KVSource.Add(nameof(propModel.部门), propModel.部门.CreateKVSource().AddRange(dataSource));
                    args.KVSource.Add(nameof(propModel.部门2), propModel.部门2.CreateKVSource().AddRange(dataSource));

                    var list = EPPlusHelper.GetList<ysbm>(args);

                    ObjectDumper.Write(list);
                    Console.WriteLine("读取完毕");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            Console.WriteLine("按任意键结束程序!");
            Console.ReadKey();
        }


        internal class ysbm
        {
            public string 序号 { get; set; }
            [KVSet("部门", true, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门 { get; set; }
            [KVSet("部门", false, "'{0}'在数据库中未找到", "部门2")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门2 { get; set; }
        }
    }
}
