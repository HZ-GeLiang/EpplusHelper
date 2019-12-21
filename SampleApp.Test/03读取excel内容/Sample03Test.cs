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

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample03Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excelList = Sample03.Run();
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample03.ExcelModel { 序号 = "1", 部门 = "娱乐1部", 部门负责人 = "赵六", 部门负责人确认签字 = "娃娃" });
            resultList.Add(new Sample03.ExcelModel { 序号 = "2", 部门 = "娱乐2部", 部门负责人 = "赵六", 部门负责人确认签字 = "菲菲" });
            resultList.Add(new Sample03.ExcelModel { 序号 = "3", 部门 = "娱乐3部", 部门负责人 = "王五", 部门负责人确认签字 = "佩琪" });
            resultList.Add(new Sample03.ExcelModel { 序号 = "4", 部门 = "娱乐4部", 部门负责人 = "jam", 部门负责人确认签字 = "jam" });
            resultList.Add(new Sample03.ExcelModel { 序号 = "5", 部门 = "娱乐5部", 部门负责人 = "静静", 部门负责人确认签字 = "静静" });
            resultList.Add(new Sample03.ExcelModel { 序号 = "6", 部门 = "娱乐6部", 部门负责人 = "jack", 部门负责人确认签字 = "jack" });
            CollectionAssert.AreEqual(excelList, resultList);

        }
    }
}
