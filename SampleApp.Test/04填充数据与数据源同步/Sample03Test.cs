using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace SampleApp.Test._04填充数据与数据源同步
{
    [TestClass]
    public class Sample03Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            SampleApp._04填充数据与数据源同步.Sample03.OpenDir = false;
            SampleApp._04填充数据与数据源同步.Sample03.Run();

            Help.GetExcelFilePath(SampleApp._04填充数据与数据源同步.Sample03.filePathSave, out var runResultFilePath, out var correctResultFilePath);

            using (var fs1 = new FileStream(correctResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var fs2 = new FileStream(runResultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage1 = new ExcelPackage(fs1))
            using (var excelPackage2 = new ExcelPackage(fs2))
            {
                Help.CompareWorkSheetCellsValue(excelPackage1, excelPackage2, 1);
            }
        }
    }
}
