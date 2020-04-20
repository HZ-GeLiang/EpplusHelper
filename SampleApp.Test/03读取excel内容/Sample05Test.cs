using EPPlusExtensions.Attributes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using System;
using System.Collections.Generic;
using System.Linq;
using EPPlusExtensions.CustomModelType;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample05Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var dataSource = new Dictionary<string, long?>();
            dataSource.Add("事业1部", 1);
            dataSource.Add("事业2部", 2);
            dataSource.Add("事业3部", null);

            var excelList = Sample05.Run(dataSource);
            var resultList = excelList.GetEmpty().ToList();
            resultList.Add(new Sample05.ExcelModel { 序号 = 1, 部门 = new KV<string, long?>("事业1部", 1), 部门2 = new KV<string, long?>("事业1部", 1) });
            resultList.Add(new Sample05.ExcelModel { 序号 = 2, 部门 = new KV<string, long?>("事业2部", 2), 部门2 = new KV<string, long?>("111") });
            resultList.Add(new Sample05.ExcelModel { 序号 = 3, 部门 = new KV<string, long?>("事业3部", null), 部门2 = new KV<string, long?>("222") });

            var index = 2;
            var a = excelList[index];
            var b = resultList[index];
            CollectionAssert.AreEqual(excelList, resultList);

        }

        [TestMethod]
        public void 数据源缺少事业2部_会异常()
        {
            var dataSource = new Dictionary<string, long?>();
            dataSource.Add("事业1部", 1);
            //dataSource.Add("事业2部", 2);
            dataSource.Add("事业3部", null);
            Assert.ThrowsException<ArgumentException>(() => Sample05.Run(dataSource));
            try
            {
                Sample05.Run(dataSource);
            }
            catch (Exception ex)
            {
                Assert.AreEqual(ex.Message, $@"'事业2部'在数据库中未找到
参数名: 部门");
            }


        }


    }
}
