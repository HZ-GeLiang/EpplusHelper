using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using OfficeOpenXml;

namespace SampleApp._03读取excel内容
{
    /// <summary>
    /// 获得模版数据检测提示.
    /// </summary>
    class Sample09
    {
        public void Run()
        {
            var wss = new List<string> { "eq", "gt", "lt", "neq1", "neq2" };
            TestMatchingModel(wss);
            Console.ReadKey();
        }

        private static void TestMatchingModel(List<string> wss)
        {
            string errMsg;
            foreach (var ws in wss)
            {
                Console.WriteLine($@"****{ws}-测试ing****");

                try
                {
                    ReadLine(ws);
                }
                catch (Exception e)
                {
                    errMsg = "模版多提供了model属性中不存在的列:C1(c)!";
                    if (ws == "gt" && e.Message != errMsg)
                    {
                        throw e;
                    }
                    errMsg = "模版少提供了model属性中定义的列:'b'!";
                    if (ws == "lt" && e.Message != errMsg)
                    {
                        throw e;
                    }
                    errMsg = "模版多提供了model属性中不存在的列:B1(c)!模版少提供了model属性中定义的列:'b'!";
                    if (ws == "neq1" && e.Message != errMsg)
                    {
                        throw e;
                    }
                    errMsg = "模版多提供了model属性中不存在的列:A1(c),B1(d)!模版少提供了model属性中定义的列:'a','b'!";
                    if (ws == "neq2" && e.Message != errMsg)
                    {
                        throw e;
                    }
                }
                Console.WriteLine($@"****{ws}-测试通过****");
            }
        }

        public static void ReadLine(string wsName)
        {
            string filePath = @"模版\03读取excel内容\Sample09.xlsx";
            using( var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                try
                {
                    var list = EPPlusHelper.GetList<Model1>(ws, 2);
                    ObjectDumper.Write(list);
                    Console.WriteLine("读取完毕");
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        class Model1
        {
            public string a { get; set; }
            public string b { get; set; }
        }
    }
}
