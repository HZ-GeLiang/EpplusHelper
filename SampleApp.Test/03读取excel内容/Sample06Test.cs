using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample06Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample06.Run<Sample06.ExcelModel>();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample06.ExcelModel { 序号 = 1, 部门 = "互娱-棋牌", 部门Id = 100, 预算部门 = "互娱-棋牌", 预算部门负责人 = "张三", 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃" });
            resultList.Add(new Sample06.ExcelModel { 序号 = 2, 部门 = "互娱-运维", 部门Id = 1002, 预算部门 = "互娱-运维", 预算部门负责人 = "李四", 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲" });
            resultList.Add(new Sample06.ExcelModel { 序号 = 3, 部门 = "审计部", 部门Id = 1003, 预算部门 = "审计部", 预算部门负责人 = "王五", 部门负责人 = "静静", 部门负责人确认签字 = "亮亮" });
            CollectionAssert.AreEqual(excelList, resultList);
        }

        [TestMethod]
        public void 修改Attribute_设置1个异常()
        {
            Assert.ThrowsException<ArgumentException>(() => Sample06.Run<Sample06.ExcelModel2>());
            try
            {
                Sample06.Run<Sample06.ExcelModel2>();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"无效的单元格:C2(部门Id:值必须在[101,99999]之间)");
                Assert.AreEqual(ex.InnerException.Message, $@"值必须在[101,99999]之间");
            }
        }

        [TestMethod]
        public void 修改Attribute_设置2个异常()
        {
            Assert.ThrowsException<ArgumentException>(() => Sample06.Run<Sample06.ExcelModel3>());
            try
            {
                Sample06.Run<Sample06.ExcelModel3>();
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"无效的单元格:B2(部门:部门名字长度要在9-10之间)");
                Assert.AreEqual(ex.InnerException.Message, $@"部门名字长度要在9-10之间");
            }
        }
    }
}
