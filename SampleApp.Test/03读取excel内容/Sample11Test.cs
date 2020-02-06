using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample11Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample11.Run().ToList();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample11.ExcelModel { 序号 = 1, 姓名 = "娃娃", JanuaryStatistics = 4, FebruaryStatistics = 7 });
            resultList.Add(new Sample11.ExcelModel { 序号 = 2, 姓名 = "菲菲", JanuaryStatistics = 5, FebruaryStatistics = 8 });
            resultList.Add(new Sample11.ExcelModel { 序号 = 3, 姓名 = "佩琪", JanuaryStatistics = 6, FebruaryStatistics = 9 });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}
