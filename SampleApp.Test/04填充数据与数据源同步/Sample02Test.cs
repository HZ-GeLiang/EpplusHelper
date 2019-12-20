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

namespace SampleApp.Test._04填充数据与数据源同步
{
    [TestClass]
    public class Sample02Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            SampleApp._04填充数据与数据源同步.Sample02.OpenDir = false;
            SampleApp._04填充数据与数据源同步.Sample02.Run();

            Help.GetExcelFilePath(SampleApp._04填充数据与数据源同步.Sample02.filePathSave, out var runResultFilePath, out var correctResultFilePath);

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
