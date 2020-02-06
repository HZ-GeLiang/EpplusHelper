using EPPlusExtensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.IO;

namespace SampleApp.Test
{
    public class Help
    {
        internal static void GetProjectPath(out string projectBinDebugPath, out string projectPath)
        {
            projectBinDebugPath = AppDomain.CurrentDomain.BaseDirectory;
            projectPath = new DirectoryInfo(projectBinDebugPath).Parent.Parent.FullName;
        }

        internal static void GetExcelFilePath(string filePath, out string runResultFilePath, out string correctResultFilePath)
        {
            Help.GetProjectPath(out var projectBinDebugPath, out var projectPath);
            runResultFilePath = Path.Combine(projectBinDebugPath, filePath);
            correctResultFilePath = Path.Combine(projectPath, filePath);
        }

        internal static void CompareWorkSheetCellsValue(ExcelPackage excelPackage1, ExcelPackage excelPackage2, int workSheetIndex1And2)
        {
            CompareWorkSheetCellsValue(excelPackage1, excelPackage2, workSheetIndex1And2, workSheetIndex1And2);
        }

        internal static void CompareWorkSheetCellsValue(ExcelPackage excelPackage1, ExcelPackage excelPackage2, int workSheetIndex1, int workSheetIndex2)
        {
            var ws1 = EPPlusHelper.GetExcelWorksheet(excelPackage1, workSheetIndex1);
            var ws2 = EPPlusHelper.GetExcelWorksheet(excelPackage2, workSheetIndex2);
            CompareWorkSheetCellsValue(ws1, ws2);
        }

        internal static void CompareWorkSheetCellsValue(ExcelPackage excelPackage1, ExcelPackage excelPackage2, string workSheetName1And2)
        {
            CompareWorkSheetCellsValue(excelPackage1, excelPackage2, workSheetName1And2, workSheetName1And2);
        }

        internal static void CompareWorkSheetCellsValue(ExcelPackage excelPackage1, ExcelPackage excelPackage2, string workSheetName1, string workSheetName2)
        {
            var ws1 = EPPlusHelper.GetExcelWorksheet(excelPackage1, workSheetName1);
            var ws2 = EPPlusHelper.GetExcelWorksheet(excelPackage2, workSheetName2);
            CompareWorkSheetCellsValue(ws1, ws2);
        }

        internal static void CompareWorkSheetCellsValue(ExcelWorksheet ws1, ExcelWorksheet ws2)
        {
            EPPlusHelper.SetSheetCellsValueFromA1(ws1);
            EPPlusHelper.SetSheetCellsValueFromA1(ws2);
            object[,] arr1 = ws1.Cells.Value as object[,];
            //Debug.Assert(arr1 != null, nameof(arr1) + " != null");
            object[,] arr2 = ws2.Cells.Value as object[,];
            CollectionAssert.AreEqual(arr1, arr2);
        }
    }
}
