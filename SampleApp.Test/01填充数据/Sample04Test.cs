﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace SampleApp.Test._01填充数据
{
    [TestClass]
    public class Sample04Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            SampleApp._01填充数据.Sample04.OpenDir = false;
            SampleApp._01填充数据.Sample04.Run();

            Help.GetExcelFilePath(SampleApp._01填充数据.Sample04.filePathSave, out var runResultFilePath, out var correctResultFilePath);

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