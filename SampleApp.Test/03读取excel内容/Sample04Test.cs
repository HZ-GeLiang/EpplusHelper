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
using EPPlusExtensions.Attributes;

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
            resultList.Add(new Sample04.ExcelModel { 序号 = "1", 部门 = new KV<string, long>("事业1部", 1) { HasValue = true }, 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃", 部门评分 = new KV<long, string>(1, "非常不满意") { HasValue = true, } });
            resultList.Add(new Sample04.ExcelModel { 序号 = "2", 部门 = new KV<string, long>("事业2部", 2) { HasValue = true }, 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲", 部门评分 = new KV<long, string>(2, "不满意") { HasValue = true, } });
            resultList.Add(new Sample04.ExcelModel { 序号 = "3", 部门 = new KV<string, long>("事业3部", 3) { HasValue = true }, 部门负责人 = "王五", 部门负责人确认签字 = "佩琪", 部门评分 = null });
            resultList.Add(new Sample04.ExcelModel { 序号 = "4", 部门 = new KV<string, long>("事业4部", 4) { HasValue = true }, 部门负责人 = "jam", 部门负责人确认签字 = "jam", 部门评分 = new KV<long, string>(3, "一般") { HasValue = true, } });
            resultList.Add(new Sample04.ExcelModel { 序号 = "6", 部门 = new KV<string, long>("事业6部", 6) { HasValue = true }, 部门负责人 = "jack", 部门负责人确认签字 = "jack", 部门评分 = new KV<long, string>(3, "一般") { HasValue = true, } });
            CollectionAssert.AreEqual(excelList, resultList);

        }
    }
}
