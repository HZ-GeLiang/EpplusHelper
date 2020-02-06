using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample08Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample08.Run().ToList();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample08.ExcelModel { 名字 = "1", 名字2 = "2", 名字3 = "3" });
            resultList.Add(new Sample08.ExcelModel { 名字 = "4", 名字2 = "5", 名字3 = "6" });

            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}
