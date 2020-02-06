using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample15Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample15.Run().ToList();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample15.ExcelModel { Num1 = 1, Num2 = 2, Sum = 3, CopyNum1 = 1, CopySum = 3 });
            resultList.Add(new Sample15.ExcelModel { Num1 = 2, Num2 = 3, Sum = 5, CopyNum1 = 2, CopySum = 5 });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}
