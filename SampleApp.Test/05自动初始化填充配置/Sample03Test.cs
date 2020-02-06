using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace SampleApp.Test._05自动初始化填充配置
{
    [TestClass]
    public class Sample03Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            SampleApp._05自动初始化填充配置.Sample03.OpenDir = false;
            SampleApp._05自动初始化填充配置.Sample03.Run();
            
            Help.GetExcelFilePath(SampleApp._05自动初始化填充配置.Sample03.FilePathSave, out var runResultFilePath, out var correctResultFilePath);

            using (var fs1 = new FileStream(correctResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var fs2 = new FileStream(runResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage1 = new ExcelPackage(fs1))
            using (var excelPackage2 = new ExcelPackage(fs2))
            {
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, "a");
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, "b");
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, "c");
            }
        }
    }
}
