using EPPlusExtensions.Validators;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SampleApp.Test._00一些方法
{
    [TestClass]
    public class ExcelSheetNameValidatorTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            {
                var validator = new ExcelSheetNameValidator("123");
                Assert.AreEqual(true, validator.IsValidSheetName());
                Assert.AreEqual("", validator.GetInvalidReason());
            }

            {
                var validator = new ExcelSheetNameValidator("1/2");
                Assert.AreEqual("工作表名称中不能包含以下字符：: \\ / ? * [ ] 或 .", validator.GetInvalidReason());
                Assert.AreEqual("1_2", validator.GetFixSheetName());
            }


        }
    }
}