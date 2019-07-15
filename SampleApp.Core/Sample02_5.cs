using System;
using System.Collections.Generic;
using System.IO;
using EPPlusExtensions;
using OfficeOpenXml;

namespace SampleApp.Core
{
    /// <summary>
    /// 获得模版数据检测提示.
    /// </summary>
    class Sample02_5
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
            string filePath = @"模版\Sample02_5.xlsx";
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                List<Model1> list;
                try
                {
                    list = EPPlusHelper.GetList<Model1>(ws, 2);
                    ObjectDumper.Write(list);
                }
                catch (Exception e)
                {
                    throw e;
                }

                Console.WriteLine("读取完毕");
            }
        }
        class Model1
        {
            public string a { get; set; }
            public string b { get; set; }
        }
    }
}
