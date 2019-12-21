using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample14Test
    {
        [TestMethod]
        public void TestMethod1()
        {

            Assert.ThrowsException<System.Exception>(() => Sample14.Run(), "数据的起始列有合并行的必须确保当前行的数据都是合并行");


        }
    }
}
