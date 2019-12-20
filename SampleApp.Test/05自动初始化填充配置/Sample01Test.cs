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

namespace SampleApp.Test._05自动初始化填充配置
{
    [TestClass]
    public class Sample01Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var txt = SampleApp._05自动初始化填充配置.Sample01.Run();
            var result = "$tb1序号	$tb1工号	$tb1姓名	$tb1性别";
            Assert.AreEqual(txt, result);
        }

    }
}
