﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System.Linq;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample01Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample01.Run();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample01.ExcelModel { 序号 = 1, 部门 = "事业1部", 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃" });
            resultList.Add(new Sample01.ExcelModel { 序号 = 2, 部门 = "事业2部", 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲" });
            resultList.Add(new Sample01.ExcelModel { 序号 = 3, 部门 = "事业3部", 部门负责人 = "王五", 部门负责人确认签字 = "佩琪" });
            resultList.Add(new Sample01.ExcelModel { 序号 = 4, 部门 = "事业4部", 部门负责人 = "jam", 部门负责人确认签字 = "jam" });
            resultList.Add(new Sample01.ExcelModel { 序号 = 4, 部门 = "事业4部", 部门负责人 = "jam", 部门负责人确认签字 = "jam" });
            resultList.Add(new Sample01.ExcelModel { 序号 = 6, 部门 = "事业6部", 部门负责人 = "jack", 部门负责人确认签字 = "jack" });
            CollectionAssert.AreEqual(excelList, resultList);
        }
    }
}