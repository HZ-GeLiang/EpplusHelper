using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample12Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample12.Run<Sample12.ExcelModel1>(true);
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample12.ExcelModel1
            {
                序号 = 1,
                姓名 = "娃娃",
                班级 = "1",
                JanuaryStatistics = 4,
                FebruaryStatistics = 7
            });
            resultList.Add(new Sample12.ExcelModel1
            {
                序号 = 2,
                姓名 = "菲菲",
                班级 = "2班",
                JanuaryStatistics = 5,
                FebruaryStatistics = 8
            });
            resultList.Add(new Sample12.ExcelModel1
            {
                序号 = 3,
                姓名 = "佩琪",
                班级 = "2班",
                JanuaryStatistics = 6,
                FebruaryStatistics = 9
            });
            CollectionAssert.AreEqual(excelList, resultList);
        }

        [TestMethod]
        public void 应该抛异常()
        {
            Assert.ThrowsException<Exception>(() => Sample12.Run<Sample12.ExcelModel2>(true));

            try
            {
                Sample12.Run<Sample12.ExcelModel2>(true);
            }
            catch (Exception ex)
            {
                /*
                Assert.AreEqual(ex.Message, $@"程序报错:Message:
无效的数字
参数名: 姓名(B列)
");
                */
                Assert.AreEqual(ex.Message, $@"程序报错:Message:
无效的数字
参数名: 姓名(B列),
无效的数字
参数名: 班级(C列)
");
            }
            try
            {
                Sample12.Run<Sample12.ExcelModel2>(false);
            }
            catch (Exception ex)
            {
                /*
                Assert.AreEqual(ex.Message, $@"程序报错:Message:
无效的数字
参数名: 姓名(B3,B4,B5)
");
                */
                Assert.AreEqual(ex.Message, $@"程序报错:Message:
无效的数字
参数名: 姓名(B3,B4,B5),
无效的数字
参数名: 班级(C4,C5)
");
            }
        }
    }
}