using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp._03读取excel内容
{
    class Sample05
    {
        public void Run()
        {
            string filePath = @"模版\03读取excel内容\Sample05.xlsx";
            var wsName = 1;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                try
                {
                    var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
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

        class ysbm
        {
            public string 序号 { get; set; }
            [KVSet("部门", true, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门 { get; set; }
            [KVSet("部门", false, "'{0}'在数据库中未找到", "部门2")]//'事业1部'在数据库中未找到
            public KV<string, long?> 部门2 { get; set; }
        }
    }
}
