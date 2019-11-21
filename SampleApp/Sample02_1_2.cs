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
    class Sample02_1_2
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_1.xlsx";
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                try
                {
                    ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<ysbm>(ws, 2);
                    args.ScanLine = ScanLine.SingleLine;

                    var propModel = new ysbm();

                    var dataSource = new Dictionary<string, long>()
                    {
                        //{"事业1部",1},
                        {"事业2部",2},
                        {"事业3部",3},
                        {"事业4部",4},
                        {"事业5部",5},
                        {"事业6部",6},
                    };

                    args.KVSource.Add(nameof(propModel.部门), propModel.部门.CreateKVSource().AddRange(dataSource));

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
            //[KVSet("部门")] // 属性'部门'值:'事业1部'未在'部门'集合中出现
            //[KVSet("部门", "部门在数据库中未找到")] //部门在数据库中未找到
            //[KVSet("部门", "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            [KVSet("部门", false, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long> 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }
        }
    }
}
