using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace SampleApp.Test._05自动初始化填充配置
{
    [TestClass]
    public class Sample02Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            SampleApp._05自动初始化填充配置.Sample02.OpenDir = false;
            SampleApp._05自动初始化填充配置.Sample02.Run();

            Help.GetProjectPath(out var projectBinDebugPath, out var projectPath);
            var fileSavePath = @"模版\05自动初始化填充配置\Sample02_Result.xlsx";
            var runResultFilePath = Path.Combine(projectBinDebugPath, fileSavePath);
            var correctResultFilePath = Path.Combine(projectPath, fileSavePath);

            using (var fs1 = new FileStream(correctResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var fs2 = new FileStream(runResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage1 = new ExcelPackage(fs1))
            using (var excelPackage2 = new ExcelPackage(fs2))
            {
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, "Sheet1");
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, "Sheet2");
            }

            var fileName1 = @"模版\05自动初始化填充配置\Sample02_Result_CrateClassSnippe_Sheet1.txt";
            var fileName2 = @"模版\05自动初始化填充配置\Sample02_Result_CrateClassSnippe_Sheet2.txt";
            var fileName3 = @"模版\05自动初始化填充配置\Sample02_Result_CrateDataTableSnippe_Sheet1.txt";
            var fileName4 = @"模版\05自动初始化填充配置\Sample02_Result_CrateDataTableSnippe_Sheet2.txt";
            var runFilePath1 = Path.Combine(projectBinDebugPath, fileName1);
            var runFilePath2 = Path.Combine(projectBinDebugPath, fileName2);
            var runFilePath3 = Path.Combine(projectBinDebugPath, fileName3);
            var runFilePath4 = Path.Combine(projectBinDebugPath, fileName4);
            var correctResultFilePath1 = Path.Combine(projectPath, fileName1);
            var correctResultFilePath2 = Path.Combine(projectPath, fileName2);
            var correctResultFilePath3 = Path.Combine(projectPath, fileName3);
            var correctResultFilePath4 = Path.Combine(projectPath, fileName4);

            Assert.AreEqual(File.ReadAllText(runFilePath1), File.ReadAllText(correctResultFilePath1));
            Assert.AreEqual(File.ReadAllText(runFilePath2), File.ReadAllText(correctResultFilePath2));
            Assert.AreEqual(File.ReadAllText(runFilePath3), File.ReadAllText(correctResultFilePath3));
            Assert.AreEqual(File.ReadAllText(runFilePath4), File.ReadAllText(correctResultFilePath4));
        }
    }
}
