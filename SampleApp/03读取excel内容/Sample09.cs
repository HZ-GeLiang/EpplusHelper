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
    public class Sample09
    {
        public static void Run()
        {
            foreach (var item in TestCaseList)
            {
                var ws = item.WsName;
                Console.WriteLine($@"****{ws}-测试ing****");
                try
                {
                    Run(ws);
                    Console.WriteLine($@"****{ws}-测试通过****");
                }
                catch (Exception ex)
                {
                    if (ex?.Message == item.ErrMsgShouldBe)
                    {
                        Console.WriteLine($@"****{ws}-测试通过****");
                    }
                    else
                    {
                        Console.WriteLine($@"****{ws}-测试不通过****{ex?.Message }");
                    }
                }
            }
        }

        public static List<TestCase> TestCaseList = new List<TestCase>
        {
            new TestCase {WsName = "eq", ErrMsgShouldBe = null},
            new TestCase {WsName = "gt", ErrMsgShouldBe = "模版多提供了model属性中不存在的列:C1(c)!"},
            new TestCase {WsName = "lt", ErrMsgShouldBe = "模版少提供了model属性中定义的列:'b'!"},
            new TestCase {WsName = "neq1", ErrMsgShouldBe = "模版多提供了model属性中不存在的列:B1(c)!模版少提供了model属性中定义的列:'b'!"},
            new TestCase {WsName = "neq2", ErrMsgShouldBe = "模版多提供了model属性中不存在的列:A1(c),B1(d)!模版少提供了model属性中定义的列:'a','b'!"},
        };

        public static List<ExcelModel> Run(string wsName)
        {
            string filePath = @"模版\03读取excel内容\Sample09.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var list = EPPlusHelper.GetList<ExcelModel>(ws, 2);
                ObjectDumper.Write(list);
                Console.WriteLine($@"{wsName}读取完毕");
                return list;
            }
        }

        public class ExcelModel
        {
            public string a { get; set; }
            public string b { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.a == y.a &&
                       this.b == y.b;
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.a.GetHashCode() +
                       this.b.GetHashCode();
            }
        }

        public class TestCase
        {
            public string WsName { get; set; }
            public string ErrMsgShouldBe { get; set; }
        }
    }
}
