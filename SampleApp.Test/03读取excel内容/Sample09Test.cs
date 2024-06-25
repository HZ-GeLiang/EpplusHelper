using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample09Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            foreach (var item in Sample09.TestCaseList)
            {
                var ws = item.WsName;
                Console.WriteLine($@"****{ws}-测试ing****");
                try
                {
                    Sample09.Run(ws);
                    if (ws == "eq")
                    {
                        Assert.AreEqual(null, item.ErrMsgShouldBe);
                    }
                    else
                    {
                        Assert.Fail("单元测试之外的");
                    }

                }
                catch (Exception ex)
                {
                    Assert.AreEqual(ex.Message, item.ErrMsgShouldBe);
                }
            }
        }
    }
}
