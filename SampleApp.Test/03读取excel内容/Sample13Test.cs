using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample13Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample13.Run();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample13.ExcelModel { A = 1, B = 1, C = 3, D = 4, E = 4, F = 6, G = 7 });
            resultList.Add(new Sample13.ExcelModel { A = 1, B = 1, C = 4, D = 4, E = 4, F = 7, G = 8 });
            resultList.Add(new Sample13.ExcelModel { A = 1, B = 1, C = 5, D = 4, E = 4, F = 8, G = 9 });
            resultList.Add(new Sample13.ExcelModel { A = 4, B = 5, C = 6, D = 7, E = 8, F = 9, G = 10 });
            resultList.Add(new Sample13.ExcelModel { A = 5, B = 6, C = 7, D = 8, E = 9, F = 10, G = 11 });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}