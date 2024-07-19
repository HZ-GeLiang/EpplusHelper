using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample07Test
    {
        [TestMethod]
        public void 未使用内置特性的异常信息()
        {
            Assert.ThrowsException<ArgumentException>(() => Sample07.Run<Sample07.ExcelModel>());
            try
            {
                Sample07.Run<Sample07.ExcelModel>();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"无效的单元格:C6");
                Assert.AreEqual(ex.InnerException.Message, $@"未找到请求的值“其他”。");
            }
        }

        [TestMethod]
        public void 使用内置特性的异常信息()
        {
            Assert.ThrowsException<ArgumentException>(() => Sample07.Run<Sample07.ExcelModel2>());
            try
            {
                Sample07.Run<Sample07.ExcelModel2>();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"无效的单元格:C6");
                Assert.AreEqual(ex.InnerException.Message, $@"张三5的性别填写不正确:'其他'");
            }
        }

        [Ignore]
        public void TestMethod1()
        {
            var excelList = Sample07.Run<Sample07.ExcelModel>();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample07.ExcelModel { 序号 = 1, 名字 = "张三1", 性别 = Sample07.Gender.男, 出生日期 = Convert.ToDateTime("2024/4/14"), 身份证号码 = "11111111111111111", 年龄 = 15 });
            resultList.Add(new Sample07.ExcelModel { 序号 = 2, 名字 = "张三2", 性别 = Sample07.Gender.女, 出生日期 = Convert.ToDateTime("1999/9/9"), 身份证号码 = "11111111111111111", 年龄 = 16 });
            resultList.Add(new Sample07.ExcelModel { 序号 = 3, 名字 = "张三3", 性别 = Sample07.Gender.未知, 出生日期 = Convert.ToDateTime("1999/9/9"), 身份证号码 = "11111111111111111", 年龄 = 17 });
            resultList.Add(new Sample07.ExcelModel { 序号 = 4, 名字 = "张三4", 性别 = null, 出生日期 = Convert.ToDateTime("1999/9/9"), 身份证号码 = "11111111111111111", 年龄 = 18 });
            resultList.Add(new Sample07.ExcelModel { 序号 = 5, 名字 = "张三5", 性别 = null, 出生日期 = Convert.ToDateTime("1999/9/9"), 身份证号码 = "11111111111111111", 年龄 = 19 });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}
