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
    public class Sample12Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample12.Run<Sample12.ExcelModel>();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample12.ExcelModel { 序号 = 1, 姓名 = "娃娃", JanuaryStatistics = 4, FebruaryStatistics = 7 });
            resultList.Add(new Sample12.ExcelModel { 序号 = 2, 姓名 = "菲菲", JanuaryStatistics = 5, FebruaryStatistics = 8 });
            resultList.Add(new Sample12.ExcelModel { 序号 = 3, 姓名 = "佩琪", JanuaryStatistics = 6, FebruaryStatistics = 9 });
            CollectionAssert.AreEqual(excelList, resultList);
        }

        [TestMethod]
        public void 应该抛异常()
        {
            Assert.ThrowsException<Exception>(() => Sample12.Run<Sample12.ExcelModel2>());

            try
            {
                Sample12.Run<Sample12.ExcelModel2>();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"程序报错:Message:无效的数字
参数名: 姓名(B列)");
            }

        }
    }
}
