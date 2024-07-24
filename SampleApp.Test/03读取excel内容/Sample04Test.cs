using EPPlusExtensions.CustomModelType;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample04Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample04.Run();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample04.ExcelModel { 序号 = "1", 部门 = new KV<string, long>("事业1部", 1), 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃", 部门评分 = new KV<long, string>(1, "非常不满意") });
            resultList.Add(new Sample04.ExcelModel { 序号 = "2", 部门 = new KV<string, long>("事业2部", 2), 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲", 部门评分 = new KV<long, string>(2, "不满意") });
            resultList.Add(new Sample04.ExcelModel { 序号 = "3", 部门 = new KV<string, long>("事业3部", 3), 部门负责人 = "王五", 部门负责人确认签字 = "佩琪", 部门评分 = null });
            resultList.Add(new Sample04.ExcelModel { 序号 = "4", 部门 = new KV<string, long>("事业4部", 4), 部门负责人 = "jam", 部门负责人确认签字 = "jam", 部门评分 = new KV<long, string>(3, "一般") });
            resultList.Add(new Sample04.ExcelModel { 序号 = "6", 部门 = new KV<string, long>("事业6部", 6), 部门负责人 = "jack", 部门负责人确认签字 = "jack", 部门评分 = new KV<long, string>(3, "一般") });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}