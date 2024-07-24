using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample14Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            Assert.ThrowsException<System.Exception>(() => Sample14.Run());
            try
            {
                Sample14.Run();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, @"检测到数据的起始列是合并行,请确保当前行的数据都是合并行.当前C2单元格不满足需求.");
            }
        }
    }
}