using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SampleApp.Test._05自动初始化填充配置
{
    [TestClass]
    public class Sample01Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var txt = SampleApp._05自动初始化填充配置.Sample01.Run();
            var result = "$tb1序号	$tb1工号	$tb1姓名	$tb1性别";
            Assert.AreEqual(txt, result);
        }

        [TestMethod]
        public void TestMethod2()
        {
            var txt = SampleApp._05自动初始化填充配置.Sample01_alias.Run();
            var result = "$tb1Index	$tb1工号	$tb1姓名	$tb1性别";
            Assert.AreEqual(txt, result);
        }
    }
}