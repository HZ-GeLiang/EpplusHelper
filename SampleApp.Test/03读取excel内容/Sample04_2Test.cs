using EPPlusExtensions.CustomModelType;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample04_2Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample04_2.Run();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample04_2.ExcelModel { 序号 = "1", 部门 = "事业1部", 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃", 部门评分 = new KV<long, string>(1, "非常不满意", "very bad") });
            resultList.Add(new Sample04_2.ExcelModel { 序号 = "2", 部门 = "事业2部", 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲", 部门评分 = new KV<long, string>(2, "不满意", "bad") });
            resultList.Add(new Sample04_2.ExcelModel { 序号 = "3", 部门 = "事业3部", 部门负责人 = "王五", 部门负责人确认签字 = "佩琪", 部门评分 = null });
            resultList.Add(new Sample04_2.ExcelModel { 序号 = "4", 部门 = "事业4部", 部门负责人 = "jam", 部门负责人确认签字 = "jam", 部门评分 = new KV<long, string>(3, "一般", "just so so") });
            resultList.Add(new Sample04_2.ExcelModel { 序号 = "6", 部门 = "事业6部", 部门负责人 = "jack", 部门负责人确认签字 = "jack", 部门评分 = new KV<long, string>(3, "一般", "just so so") });

            var index = 0;
            var a = excelList[index];
            var b = resultList[index];
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}