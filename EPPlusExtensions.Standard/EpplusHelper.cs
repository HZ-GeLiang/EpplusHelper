﻿using EPPlusExtensions.Attributes;
using EPPlusExtensions.CustomModelType;
using EPPlusExtensions.Exceptions;
using EPPlusExtensions.ExtensionMethods;
using EPPlusExtensions.Helpers;
using EPPlusExtensions.Validators;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusExtensions
{
    public class EPPlusHelper
    {
        /// <summary>
        /// 填充Excel时创建的工作簿名字
        /// </summary>
        public static List<string> FillDataWorkSheetNameList = new List<string>();

        public const string XlsxContentType = ContentTypes.XlsxContentType;

        #region GetExcelWorksheet

        /// <summary>
        /// 获得当前Excel的所有工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelWorksheet> GetExcelWorksheets(ExcelPackage excelPackage)
        {
            for (var i = 1; i <= excelPackage.Workbook.Worksheets.Count; i++)
            {
                //    var ws = excelPackage.Compatibility.IsWorksheets1Based
                //        ? excelPackage.Workbook.Worksheets[i]
                //        : excelPackage.Workbook.Worksheets[i - 1];
                //    yield return ws;
                yield return GetExcelWorksheet(excelPackage, i);
            }
        }

        /// <summary>
        /// 获得Excel的第N个Sheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始</param>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, int workSheetIndex)
        {
            if (workSheetIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            }
            workSheetIndex = ExcelPackageHelper.ConvertWorkSheetIndex(excelPackage, workSheetIndex);
            var ws = excelPackage.Workbook.Worksheets[workSheetIndex];
            return ws;
        }

        /// <summary>
        /// 根据workSheetIndex获得模版worksheet,然后复制一份并重命名成workSheetName后返回
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="copyWorkSheetIndex">从1开始</param>
        /// <param name="workSheetNewName"></param>
        /// <returns></returns>
        public static ExcelWorksheet DuplicateWorkSheetAndRename(ExcelPackage excelPackage, int copyWorkSheetIndex, string workSheetNewName)
        {
            if (copyWorkSheetIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(copyWorkSheetIndex));
            }
            if (workSheetNewName is null)
            {
                throw new ArgumentNullException(nameof(workSheetNewName));
            }
            //您为工作表或图表输入的名称无效。请确保：
            //    ·名称不多于31个字符。
            //    ·名称不包含下列任一字符:：\/？*[或]。   注意, 对于： 只有全角和半角字符, 但是这2个都不可以
            //    ·名称不为空。
            if (workSheetNewName.Length > 31)
            {
                throw new ArgumentNullException(nameof(workSheetNewName) + "名称不多于31个字符");
            }
            var violateChars = new[] { ':', '：', '\\', '/', '？', '*', '[', ']' };
            if (violateChars.Any(violateChar => workSheetNewName.Contains(violateChar)))
            {
                throw new ArgumentNullException(nameof(workSheetNewName) + "名称不包含下列任一字符:：\\/？*[或]。");
            }
            if (workSheetNewName.Length <= 0)
            {
                throw new ArgumentNullException(nameof(workSheetNewName) + "名称不为空");
            }

            var wsMom = GetExcelWorksheet(excelPackage, copyWorkSheetIndex);
            var ws = excelPackage.Workbook.Worksheets.Add(workSheetNewName, wsMom);
            ws.Name = new ExcelSheetNameValidator(workSheetNewName).GetFixSheetName();
            return ws;
        }

        /// <summary>
        /// 根据worksheet名字获得worksheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workName"></param>
        /// <returns></returns>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string workName)
        {
            return GetExcelWorksheet(excelPackage, workName, false);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workName"></param>
        /// <param name="onlyOneWorkSheetReturnThis">用于判断是否在 Excel 文档仅有一个工作表时返回该工作表</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string workName, bool onlyOneWorkSheetReturnThis)
        {
            ExcelWorksheet ws = null;
            if (onlyOneWorkSheetReturnThis && excelPackage.Workbook.Worksheets.Count == 1)
            {
                var workSheetIndex = ExcelPackageHelper.ConvertWorkSheetIndex(excelPackage, 1);
                ws = excelPackage.Workbook.Worksheets[workSheetIndex];
                if (ws != null)
                {
                    return ws;
                }
            }
            if (workName is null)
            {
                throw new ArgumentNullException(nameof(workName));
            }

            ws = excelPackage.Workbook.Worksheets[workName];
            if (ws != null)
            {
                return ws;
            }
            throw new ArgumentException($@"当前Excel中不存在名为'{workName}'的worksheet", nameof(workName));
        }

        /// <summary>
        /// 获得当前Excel的所有工作簿的名字
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        public static List<string> GetExcelWorksheetNames(ExcelPackage excelPackage)
        {
            //for (int i = 1; i <= excelPackage.Workbook.Worksheets.Count; i++)
            //{
            //    wsNameList.Add(GetExcelWorksheet(excelPackage, i).Name);
            //}
            //return wsNameList;
            var wsNameList = GetExcelWorksheets(excelPackage).Select(item => item.Name).ToList();
            return wsNameList;
        }

        /// <summary>
        /// 根据名字获取Worksheet,然后复制一份出来并重命名成workSheetName并返回
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="destWorkSheetName">填充数据的workSheet叫什么</param>
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么</param>
        /// <returns></returns>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string destWorkSheetName, string workSheetNewName)
        {
            if (destWorkSheetName is null)
            {
                throw new ArgumentNullException(nameof(destWorkSheetName));
            }
            if (workSheetNewName is null)
            {
                throw new ArgumentNullException(nameof(workSheetNewName));
            }

            var wsTemplate = GetExcelWorksheet(excelPackage, destWorkSheetName);
            try
            {
                var ws = excelPackage.Workbook.Worksheets.Add(workSheetNewName, wsTemplate);
                ws.Name = new ExcelSheetNameValidator(workSheetNewName).GetFixSheetName();
                return ws;
            }
            catch (NullReferenceException ex)
            {
                //遇到场景记录:读取worksheet , 然后用代码创建配置信息, 然后在调用Fill时, 遇到了这个错误
                throw new Exception($"受Epplus的限制, 无法复制'{workSheetNewName}'工作簿", ex);
            }
        }

        #endregion

        #region DeleteWorksheet

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetName"></param>
        public static void DeleteWorksheet(ExcelPackage excelPackage, string workSheetName)
        {
            if (workSheetName is null)
            {
                throw new ArgumentNullException(nameof(workSheetName));
            }

            if (excelPackage.Workbook.Worksheets.Count <= 1) //The workbook must contain at least one worksheet
            {
                return;
            }
            if (excelPackage.Workbook.Worksheets[workSheetName] != null)
            {
                excelPackage.Workbook.Worksheets.Delete(workSheetName);
            }
        }

        /// <summary>
        ///  尝试删除,如果删除的目标不存在,也不会报错
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始,注:每删除一个ws后,索引重新计算</param>
        public static void DeleteWorksheet(ExcelPackage excelPackage, int workSheetIndex)
        {
            if (excelPackage.Workbook.Worksheets.Count <= 1) //The workbook must contain at least one worksheet
            {
                return;
            }

            workSheetIndex = ExcelPackageHelper.ConvertWorkSheetIndex(excelPackage, workSheetIndex);
            var ws = excelPackage.Workbook.Worksheets[workSheetIndex];
            if (ws != null)
            {
                excelPackage.Workbook.Worksheets.Delete(workSheetIndex);
            }
        }

        /// <summary>
        /// 删除工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="eWorkSheetHiddens">获得工作簿的参数</param>
        public static void DeleteWorksheet(ExcelPackage excelPackage, params eWorkSheetHidden[] eWorkSheetHiddens)
        {
            EPPlusHelper.DeleteWorksheet(excelPackage, new List<string>(), eWorkSheetHiddens);
        }

        /// <summary>
        /// 删除工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetNameExcludes">要保留的工作簿名字</param>
        /// <param name="eWorkSheetHiddens">获得工作簿的参数</param>
        public static void DeleteWorksheet(ExcelPackage excelPackage, string[] workSheetNameExcludes, params eWorkSheetHidden[] eWorkSheetHiddens)
        {
            EPPlusHelper.DeleteWorksheet(excelPackage, (workSheetNameExcludes ?? new string[] { }).ToList(), eWorkSheetHiddens);
        }

        /// <summary>
        /// 删除工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetNameExcludeList">要保留的工作簿名字</param>
        /// <param name="eWorkSheetHiddens">获得工作簿的参数</param>
        public static void DeleteWorksheet(ExcelPackage excelPackage, List<string> workSheetNameExcludeList, params eWorkSheetHidden[] eWorkSheetHiddens)
        {
            if (eWorkSheetHiddens is null)
            {
                return;
            }

            if (workSheetNameExcludeList is null)
            {
                workSheetNameExcludeList = new List<string>();
            }

            var delWsNames = GetWorkSheetNames(excelPackage, eWorkSheetHiddens);
            foreach (var wsName in delWsNames)
            {
                if (workSheetNameExcludeList.Contains(wsName))
                {
                    continue;
                }

                EPPlusHelper.DeleteWorksheet(excelPackage, wsName);
            }
        }

        /// <summary>
        /// 获得excel有哪些工作簿名称
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="eWorkSheetHiddens"> 可以用来获得隐藏的工作簿</param>
        /// <returns></returns>
        public static List<string> GetWorkSheetNames(ExcelPackage excelPackage, params eWorkSheetHidden[] eWorkSheetHiddens)
        {
            var wsNames = new List<string>();
            if (eWorkSheetHiddens is null || eWorkSheetHiddens.Length == 0)
            {
                return wsNames;
            }

            for (int i = 1; i <= excelPackage.Workbook.Worksheets.Count; i++)
            {
                var index = ExcelPackageHelper.ConvertWorkSheetIndex(excelPackage, i);
                var ws = excelPackage.Workbook.Worksheets[index];
                //foreach (var eWorkSheetHidden in eWorkSheetHiddens)
                //{
                //    if (ws.Hidden != eWorkSheetHidden) continue;
                //    wsNames.Add(ws.Name);
                //}//代码优化
                wsNames.AddRange(from eWorkSheetHidden in eWorkSheetHiddens where ws.Hidden == eWorkSheetHidden select ws.Name);
            }
            return wsNames;
        }

        /// <summary>
        /// 删除所有的工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetNameExclude">排除的工作簿名字列表</param>
        public static void DeleteWorksheetAll(ExcelPackage excelPackage, params string[] workSheetNameExclude)
        {
            EPPlusHelper.DeleteWorksheet(excelPackage, (workSheetNameExclude ?? new string[] { }).ToList());
        }

        /// <summary>
        /// 删除所有的工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetNameExcludeList">排除的工作簿名字列表</param>
        public static void DeleteWorkSheetAll(ExcelPackage excelPackage, List<string> workSheetNameExcludeList)
        {
            EPPlusHelper.DeleteWorksheet(excelPackage, workSheetNameExcludeList ?? new List<string>(),
                eWorkSheetHidden.Hidden, eWorkSheetHidden.VeryHidden, eWorkSheetHidden.Visible);
        }

        #endregion

        #region FillData

        /// <summary>
        /// 往目标sheet中填充数据.注:填充的数据的类型会影响你设置的excel单元格的格式是否起作用
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么</param>
        /// <param name="destWorkSheetName">填充数据的workSheet叫什么</param>
        public static void FillData(ExcelPackage excelPackage, EPPlusConfig config, EPPlusConfigSource configSource, string workSheetNewName, string destWorkSheetName)
        {
            if (workSheetNewName is null)
            {
                throw new ArgumentNullException(nameof(workSheetNewName));
            }
            if (destWorkSheetName is null)
            {
                throw new ArgumentNullException(nameof(destWorkSheetName));
            }
            ExcelWorksheet worksheet = GetExcelWorksheet(excelPackage, destWorkSheetName, workSheetNewName);
            EPPlusHelper.FillData(config, configSource, worksheet);
        }

        /// <summary>
        /// 往目标sheet中填充数据.注:填充的数据的类型会影响你设置的excel单元格的格式是否起作用
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么. 若为""/null,则默认是"Sheet" +workSheetNewName </param>
        /// <param name="destWorkSheetIndex">填充数据的workSheet, 从1开始</param>
        public static void FillData(ExcelPackage excelPackage, EPPlusConfig config, EPPlusConfigSource configSource, string workSheetNewName, int destWorkSheetIndex)
        {
            if (workSheetNewName is null)
            {
                throw new ArgumentNullException(nameof(workSheetNewName));
            }
            if (destWorkSheetIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(destWorkSheetIndex));
            }

            ExcelWorksheet worksheet = EPPlusHelper.DuplicateWorkSheetAndRename(excelPackage, destWorkSheetIndex, workSheetNewName);

            EPPlusHelper.FillData(config, configSource, worksheet);
        }

        /// <summary>
        /// 往目标sheet中填充数据
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        public static void FillData(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            EPPlusHelper.FillDataWorkSheetNameList.Add(worksheet.Name);//往list里添加数据
            config.WorkSheetDefault?.Invoke(worksheet);

            EPPlusConfigHelper.FillData_Head(config, configSource, worksheet);
            int sheetBodyAddRowCount = 0;
            if (configSource?.Body?.ConfigList.Count > 0)
            {
                int allDataTableRows = 0;
                foreach (var bodyInfo in configSource.Body.ConfigList)
                {
                    allDataTableRows += bodyInfo?.Option?.DataSource?.Rows.Count ?? 0;
                }
                //这个限制仅仅针对,标题行是单行, 且填充数据是单行,没有FillData_Head 和 FillData_Foot 才有效
                if (allDataTableRows > EPPlusConfig.MaxRow07 - 1)//-1是去掉第一行的标题行
                {
                    throw new IndexOutOfRangeException("要导出的数据行数超过excel最大行限制");
                }

                sheetBodyAddRowCount = EPPlusConfigHelper.FillData_Body(config, configSource, worksheet);
            }

            EPPlusConfigHelper.FillData_Foot(config, configSource, worksheet, sheetBodyAddRowCount);
        }

        #endregion

        #region FillExcelDefaultConfig

        /// <summary>
        ///
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileOutDirectoryName"></param>
        /// <param name="dataConfigInfo"></param>
        /// <param name="cellCustom">对单元格进行额外处理</param>
        /// <returns></returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(string filePath, string fileOutDirectoryName, List<ExcelDataConfigInfo> dataConfigInfo, Action<ExcelRange> cellCustom = null)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var defaultConfigList = FillExcelDefaultConfig(excelPackage, dataConfigInfo, cellCustom);

                var haveConfig = defaultConfigList.Find(a => a.ClassPropertyList.Count > 0) != null;
                if (haveConfig)
                {
                    var path = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result.xlsx";
                    EPPlusHelper.Save(excelPackage, path);
                }

                return defaultConfigList;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="dataConfigInfo">指定的worksheet</param>
        /// <param name="cellCustom">对单元格进行额外处理</param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, List<ExcelDataConfigInfo> dataConfigInfo, Action<ExcelRange> cellCustom = null)
        {
            if (dataConfigInfo != null)
            {
                foreach (var item in dataConfigInfo)
                {
                    //WorkSheetIndex没设置,但是设置了WorkSheetName
                    if (!string.IsNullOrEmpty(item.WorkSheetName) || item.WorkSheetIndex <= 0)
                    {
                        continue;
                    }

                    var eachCount = 1;
                    foreach (var ws in excelPackage.Workbook.Worksheets)
                    {
                        if (item.WorkSheetIndex == eachCount)
                        {
                            item.WorkSheetName = ws.Name;
                            break;
                        }
                        eachCount++;
                    }
                }
            }

            var list = new List<DefaultConfig>();
            foreach (var ws in excelPackage.Workbook.Worksheets)
            {
                int titleLine = 1;
                int titleColumn = 1;
                if (dataConfigInfo is null)
                {
                    list.Add(FillExcelDefaultConfig(ws, titleLine, titleColumn, cellCustom));
                    continue;
                }

                var configInfo = dataConfigInfo.Find(a => a.WorkSheetName == ws.Name);
                if (configInfo is null)
                {
                    continue;
                }

                if (configInfo.TitleLine == 0 && configInfo.TitleColumn == 0)
                {
                    continue;
                }
                var address = ExcelWorksheetHelper.GetMergeCellAddressPrecise(ws, row: configInfo.TitleLine, col: configInfo.TitleColumn);
                var cellRange = new ExcelCellRange(address);
                if (cellRange.IsMerge)
                {
                    titleLine = cellRange.End.Row;
                    titleColumn = cellRange.End.Col;
                }
                else
                {
                    titleLine = cellRange.Start.Row;
                    titleColumn = cellRange.Start.Col;
                }
                list.Add(FillExcelDefaultConfig(ws, titleLine, titleColumn, cellCustom));
                continue;
            }
            return list;
        }

        /// <summary>
        /// 填充excel默认配置
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="titleLineNumber"></param>
        /// <param name="titleColumnNumber"></param>
        /// <param name="cellCustom">对单元格进行额外处理</param>
        /// <returns></returns>
        public static DefaultConfig FillExcelDefaultConfig(ExcelWorksheet ws, int titleLineNumber, int titleColumnNumber, Action<ExcelRange> cellCustom = null)
        {
            var colNameList = new List<ExcelCellInfoValue>();
            var nameRepeatCounter = new Dictionary<string, int>();

            string mergeCellAddress;
            #region 获得colNameList

            int col = titleColumnNumber;
            while (col <= EPPlusConfig.MaxCol07)
            {
                var excelColName = ws.Cells[titleLineNumber, col].Merge
                    ? ExcelWorksheetHelper.GetMergeCellText(ws, titleLineNumber, col)
                    : ExcelWorksheetHelper.GetCellText(ws, titleLineNumber, col);

                var destColVal = NamingHelper.ExtractName(excelColName).Trim().MergeLines();
                if (string.IsNullOrEmpty(destColVal))
                {
                    break;
                }

                var thisColNameInfo = new ExcelCellInfoValue()
                {
                    Name = destColVal,
                    ExcelColNameIndex = col,
                    //ExcelColName = excelColName,
                    ExcelColName = excelColName.Trim().MergeLines(),
                };

                NamingHelper.AutoRename(colNameList, nameRepeatCounter, thisColNameInfo, true);

                if (ExcelWorksheetHelper.IsMergeCell(ws, titleLineNumber, col, out mergeCellAddress))
                {
                    var range = new ExcelCellRange(mergeCellAddress);
                    col += range.IntervalCol + 1;
                }
                else
                {
                    col++;
                }
            }

            #endregion

            if (colNameList.Count == 0)
            {
                return new DefaultConfig()
                {
                    WorkSheetName = ws.Name,
                    CrateDataTableSnippe = null,
                    CrateClassSnippe = null,
                    ClassPropertyList = colNameList
                };
            }

            #region 给单元格赋值

            int fillBodyLine;
            if (ExcelWorksheetHelper.IsMergeCell(ws, titleLineNumber, 1, out mergeCellAddress))
            {
                var range = new ExcelCellRange(mergeCellAddress);
                fillBodyLine = range.Start.Row + range.IntervalRow + 1;
            }
            else
            {
                fillBodyLine = titleLineNumber + 1;
            }

            for (int i = 0; i < colNameList.Count; i++)
            {
                ExcelRange cell = ws.Cells[fillBodyLine, colNameList[i].ExcelColNameIndex];
                string cellValue = $"$tb1{(colNameList[i].IsRename ? colNameList[i].NameNew : colNameList[i].Name)}";
                ExcelRangeHelper.SetWorksheetCellValue(cell, cellValue);
                cellCustom?.Invoke(ws.Cells[titleLineNumber + 1, i + 1]);
            }

            #endregion

            #region sb_CreateClassSnippet + sb_CreateDateTableSnippet

            StringBuilder sb_CreateClassSnippet = new StringBuilder();
            sb_CreateClassSnippet.AppendLine($"public class {ws.Name} {{");

            StringBuilder sb_CreateDateTableSnippet = new StringBuilder();
            sb_CreateDateTableSnippet.AppendLine($@"DataTable dt = new DataTable();");
            StringBuilder sbColumn = new StringBuilder();
            StringBuilder sbAddDr = new StringBuilder();
            StringBuilder sbColumnType = new StringBuilder();
            sbAddDr.AppendLine($@"//var dr = dt.NewRow();");

            foreach (var colName in colNameList)
            {
                var propName = colName.IsRename ? colName.NameNew : colName.Name;
                sbColumn.AppendLine($"dt.Columns.Add(\"{propName}\");");
                sbAddDr.AppendLine($"//dr[\"{propName}\"] = ");
                bool sb_CrateClassSnippe_AppendLine_InForeach = false;

                if (colName.IsRename)
                {
                    sb_CreateClassSnippet.AppendLine($" [ExcelColumnIndex({colName.ExcelColNameIndex})]");
                    sb_CreateClassSnippet.AppendLine($" [DisplayExcelColumnName(\"{colName.ExcelColName}\")]");
                }

                if (colName.ExcelColName != colName.Name)
                {
                    if (!colName.IsRename)//上面添加过了,这里不在添加
                    {
                        sb_CreateClassSnippet.AppendLine($" [DisplayExcelColumnName(\"{colName.ExcelColName}\")]");
                    }
                }

                if (EpplusHelperConfig.KeysTypeOfDateTime.Any(item => propName.Contains(item)))
                {
                    sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(DateTime);");
                    sb_CreateClassSnippet.AppendLine($" public DateTime {propName} {{ get; set; }}");
                    sb_CrateClassSnippe_AppendLine_InForeach = true;
                }

                if (EpplusHelperConfig.KeysTypeOfString.Any(item => propName.Contains(item)))
                {
                    sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(string);");
                    sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                    sb_CrateClassSnippe_AppendLine_InForeach = true;
                }

                if (EpplusHelperConfig.KeysTypeOfDecimal.Any(item => propName.Contains(item)))
                {
                    sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(decimal);");
                    sb_CreateClassSnippet.AppendLine($" public decimal {propName} {{ get; set; }}");
                    sb_CrateClassSnippe_AppendLine_InForeach = true;
                }

                if (!sb_CrateClassSnippe_AppendLine_InForeach)
                {
                    sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                }
            }
            sb_CreateDateTableSnippet.Append(sbColumn);
            sb_CreateDateTableSnippet.Append(sbColumnType);
            sbAddDr.AppendLine("//dt.Rows.Add(dr);");
            sb_CreateDateTableSnippet.Append(sbAddDr);

            sb_CreateClassSnippet.AppendLine("}");
            #endregion

            return new DefaultConfig()
            {
                WorkSheetName = ws.Name,
                CrateDataTableSnippe = sb_CreateDateTableSnippet.ToString(),
                CrateClassSnippe = sb_CreateClassSnippet.ToString(),
                ClassPropertyList = colNameList
            };
        }

        #endregion

        #region SetConfigSource xxx (Head与foot配置数据源)

        /// <summary>
        /// 设置Head配置的数据源
        /// </summary>
        /// <param name="configSource"></param>
        /// <param name="dt">用来获得列名</param>
        public static void SetConfigSourceHead(EPPlusConfigSource configSource, DataTable dt)
        {
            EPPlusHelper.SetConfigSourceHead(configSource, dt, dt.Rows[0]);
        }

        /// <summary>
        /// 设置Head配置的数据源
        /// </summary>
        /// <param name="configSource"></param>
        /// <param name="dt">用来获得列名</param>
        /// <param name="dr">数据源是这个</param>
        public static void SetConfigSourceHead(EPPlusConfigSource configSource, DataTable dt, DataRow dr)
        {
            //var dict = new Dictionary<string, string>();
            //for (int i = 0; i < dr.ItemArray.Length; i++)
            //{
            //    var colName = dt.Columns[i].ColumnName;
            //    if (!dict.ContainsKey(colName))
            //    {
            //        dict.Add(colName, dr[i] == DBNull.Value || dr[i] is null ? "" : dr[i].ToString());
            //    }
            //    else
            //    {
            //        throw new Exception(nameof(SetConfigSourceHead) + "方法异常");
            //    }
            //}

            //var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell>();
            //foreach (var item in dict)
            //{
            //    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key, FillValue = dict.Values });
            //}

            //configSource.Head = new EPPlusConfigSourceHead() { CellsInfoList = fixedCellsInfoList };

            configSource.Head = new EPPlusConfigSourceHead() { CellsInfoList = EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dt, dr) };
        }

        /// <summary>
        /// 设置Foot配置的数据源
        /// </summary>
        /// <param name="configSource"></param>
        /// <param name="dt">用来获得列名</param>
        public static void SetConfigSourceFoot(EPPlusConfigSource configSource, DataTable dt)
        {
            SetConfigSourceFoot(configSource, dt, dt.Rows[0]);
        }

        /// <summary>
        /// 设置Foot配置的数据源
        /// </summary>
        /// <param name="configSource"></param>
        /// <param name="dt">用来获得列名</param>
        /// <param name="dr">数据源是这个</param>
        public static void SetConfigSourceFoot(EPPlusConfigSource configSource, DataTable dt, DataRow dr)
        {
            //var dict = new Dictionary<string, string>();
            //for (int i = 0; i < dr.ItemArray.Length; i++)
            //{
            //    var colName = dt.Columns[i].ColumnName;
            //    if (!dict.ContainsKey(colName))
            //    {
            //        dict.Add(colName, dr[i] == DBNull.Value || dr[i] is null ? "" : dr[i].ToString());
            //    }
            //    else
            //    {
            //        throw new Exception(nameof(SetConfigSourceFoot) + "方法异常");
            //    }
            //}

            //var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell>();
            //foreach (var item in dict)
            //{
            //    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key, FillValue = dict.Values });
            //}

            //configSource.Foot = new EPPlusConfigSourceFoot { CellsInfoList = fixedCellsInfoList };
            configSource.Foot = new EPPlusConfigSourceFoot { CellsInfoList = EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dt, dr) };
        }

        #endregion

        #region GetList<T>

        private static List<PropertyInfo> ICustomersModelTypeList;

        /// <summary>
        /// 初始化参数模型
        /// 无法添加 new() 约束, 因为 datarow 就是没有的
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static T InitGetExcelListArgsModel<T>() where T : class/*,new()*/
        {
            ICustomersModelTypeList = new List<PropertyInfo>();

            var t_ctor = typeof(T).GetConstructor(new Type[] { });
            if (t_ctor is null)
            {
                return default(T);
            }

            var model = t_ctor.Invoke(new object[] { });

            foreach (PropertyInfo p in ReflectionHelper.GetProperties(typeof(T)))
            {
                if (p.PropertyType.IsValueType)
                {
                    continue;
                }

                var p_ctor = p.PropertyType.GetConstructor(new Type[] { });
                if (p_ctor is null)
                {
                    continue;
                }

                p.SetValue(model, p_ctor.Invoke(new object[] { }));

                if (typeof(ICustomersModelType).IsAssignableFrom(p.PropertyType))
                {
                    ICustomersModelTypeList.Add(p);
                }
            }
            return (T)model;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="rowStart"><inheritdoc cref="GetExcelListArgs.DataRowStart" path="/summary"/></param>
        /// <returns></returns>
        public static GetExcelListArgs GetExcelListArgsDefault(ExcelWorksheet ws, int rowStart)
        {
            var args = new GetExcelListArgs
            {
                ws = ws,
                DataRowStart = rowStart,
                DataTitleRow = rowStart - 1,
                EveryCellPrefix = "",
                EveryCellReplaceList = null,
                UseEveryCellReplace = true,
                ReadCellValueOption = ReadCellValueOption.Trim,
                POCO_Property_AutoRename_WhenRepeat = false,
                POCO_Property_AutoRenameFirtName_WhenRepeat = true,
                ScanLine = ScanLine.MergeLine,
                MatchingModelEqualsCheck = true,
                GetList_NeedAllException = false,
                GetList_ErrorMessage_OnlyShowColomn = false,
                DataColStart = 1,
                DataColEnd = EPPlusConfig.MaxCol07,
            };
            return args;
        }

        /// <summary>
        ///
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowStart"><inheritdoc cref="GetExcelListArgs.DataRowStart" path="/summary"/></param>
        /// <returns></returns>
        public static GetExcelListArgs<T> GetExcelListArgsDefault<T>(ExcelWorksheet ws, int rowStart) where T : class
        {
            //这3个属性的 <T> 版本多出来的, 其余的默认值调用 GetExcelListArgsDefault(),然后用反射赋值
            var argsReturn = new GetExcelListArgs<T>
            {
                HavingFilter = null,
                WhereFilter = null,
                Model = InitGetExcelListArgsModel<T>(),
            };
            var args = GetExcelListArgsDefault(ws, rowStart);
            var dict = ReflectionHelper.GetProperties(typeof(GetExcelListArgs))
                        .ToDictionary(item => item.Name, item => item.GetValue(args));

            #region 反射赋值

            foreach (var item in ReflectionHelper.GetProperties(typeof(GetExcelListArgs<T>)))
            {
                if (dict.ContainsKey(item.Name))
                {
                    item.SetValue(argsReturn, dict[item.Name]);
                }
            }
            #endregion

            return argsReturn;
        }

        /// <summary>
        /// 只能是最普通的excel.(每个单元格都是未合并的,第一行是列名,数据从第一列开始填充的那种.)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowStart">数据起始行(不含列名),从1开始</param>
        /// <returns></returns>
        public static IEnumerable<T> GetList<T>(ExcelWorksheet ws, int rowStart) where T : class, new()
        {
            var args = EPPlusHelper.GetExcelListArgsDefault<T>(ws, rowStart);
            return EPPlusHelper.GetList<T>(args);
        }

        /// <summary>
        /// 只能是最普通的excel.(第一行是必须是列名,数据填充列起始必须是A2单元格,且每个单元格都是未合并的)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowStart">数据起始行(不含列名),从1开始</param>
        /// <param name="everyCellReplaceOldValue"></param>
        /// <param name="everyCellReplaceNewValue"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetList<T>(ExcelWorksheet ws, int rowStart, string everyCellReplaceOldValue, string everyCellReplaceNewValue) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowStart);
            if (everyCellReplaceOldValue != null && everyCellReplaceNewValue != null)
            {
                args.EveryCellReplaceList = new Dictionary<string, string> { { everyCellReplaceOldValue, everyCellReplaceNewValue } };
            }
            return GetList<T>(args);
        }

        public static IEnumerable<T> GetList<T>(ExcelWorksheet ws, int rowStart, string everyCellPrefix, Dictionary<string, string> everyCellReplace) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowStart);
            args.EveryCellPrefix = everyCellPrefix;
            args.EveryCellReplaceList = everyCellReplace;
            return GetList<T>(args);
        }

        public static IEnumerable<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {
            string GetValue(PropertyInfo pInfo, int rowStart, int colIndex)
            {
                //string value;
                //if (pInfo.PropertyType == typeof(DateTime?) || pInfo.PropertyType == typeof(DateTime))
                //{
                //    //todo:对于日期类型的,有时候要获取Cell.Value, 有空了修改
                //    value = GetMergeCellText(args.ws, rowStart, colIndex);
                //}
                //else
                //{
                //    value = GetMergeCellText(args.ws, rowStart, colIndex);
                //}
                //return value;
                return ExcelWorksheetHelper.GetMergeCellText(args.ws, rowStart, colIndex);
            }

            var colNameList = GetExcelColumnOfModel(args);//主要是计算DataColEnd的值, 放在第一行还是因为 单元测试 03.02的示例

            void Check()
            {
                if (args.DataRowStart <= 0)
                {
                    throw new ArgumentException($@"数据起始行值'{args.DataRowStart}'错误,值应该大于0");
                }

                if (args.DataTitleRow <= 0)
                {
                    throw new ArgumentException($@"数据起始行的标题行值'{args.DataTitleRow}'错误,值应该大于0");
                }

                #region 对ScanLine.MergeLine进行模版验证

                if (args.ScanLine == ScanLine.MergeLine)
                {
                    SetSheetCellsValueFromA1(args.ws);
                    object[,] arr = args.ws.Cells.Value as object[,];

                    for (int i = 0; i < arr.GetLength(0);) //遍历行,这里 i++ 没有写
                    {
                        var rowNo = i + 1;
                        if (rowNo < args.DataRowStart)
                        {
                            i++;
                            continue;
                        }

                        //如果数据的第一列的合并单元格,必须确保这一行的所有列都是合并单元格
                        if (!ExcelWorksheetHelper.IsMergeCell(args.ws, row: rowNo, col: args.DataColStart, out var mergeCellAddress))
                        {
                            i++;
                            continue;
                        }
                        else
                        {
                            i += new ExcelAddress(mergeCellAddress).Rows; //按第一列合并的行数进行step的增加
                        }

                        for (int j = 0; j < arr.GetLength(1); j++) //遍历列
                        {
                            var colNo = j + 1;
                            if (colNo < args.DataColStart || colNo > args.DataColEnd)
                            {
                                continue;
                            }

                            if (!ExcelWorksheetHelper.IsMergeCell(args.ws, row: rowNo, col: colNo))
                            {
                                throw new Exception($@"检测到数据的起始列是合并行,请确保当前行的数据都是合并行.当前{new ExcelCellPoint(rowNo, colNo).R1C1}单元格不满足需求."); //参考 示例 03.14  或03.02的模版
                            }
                        }
                    }
                }

                #endregion
            }

            Check();

            Type type = typeof(T);

            #region 获得字典

            //1.初始化3个dict
            var dictModelPropNameExistsExcelColumn = new Dictionary<string, bool>();//Model属性在Excel列中存在, key: ModelPropName
            var dictModelPropNameToExcelColumnName = new Dictionary<string, string>();//Model属性名字对应的excel的标题列名字
            var dictExcelColumnIndexToModelPropName_Temp = new Dictionary<int, string>();//Excel的列标题和Model属性名字的映射

            foreach (var propInfo in type.GetProperties())
            {
                if (ReflectionHelper.GetAttributeForProperty<IngoreAttribute>(type, propInfo.Name).Any())
                {
                    continue;
                }

                dictModelPropNameExistsExcelColumn.Add(propInfo.Name, false);
                dictModelPropNameToExcelColumnName.Add(propInfo.Name, null);

                var propAttr_DisplayExcelColumnName = ReflectionHelper.GetAttributeForProperty<DisplayExcelColumnNameAttribute>(type, propInfo.Name);
                if (propAttr_DisplayExcelColumnName.Length > 0)
                {
                    dictModelPropNameToExcelColumnName[propInfo.Name] = ((DisplayExcelColumnNameAttribute)propAttr_DisplayExcelColumnName[0]).Name;
                }
                var propAttr_ExcelColumnIndex = ReflectionHelper.GetAttributeForProperty<ExcelColumnIndexAttribute>(type, propInfo.Name);
                if (propAttr_ExcelColumnIndex.Length > 0)
                {
                    dictExcelColumnIndexToModelPropName_Temp.Add(((ExcelColumnIndexAttribute)propAttr_ExcelColumnIndex[0]).Index, propInfo.Name);
                }
            }

            var dictExcelColumnIndexToModelPropName_All = new Dictionary<int, string>();//Excel列对应的Model属性名字(所有excel列)
            var dictExcelAddressCol = colNameList.ToDictionary(item => item.ExcelAddress, item => new ExcelCellPoint(item.ExcelAddress).Col);
            //初始化 dictExcelColumnIndexToModelPropName_All
            foreach (var item in colNameList)
            {
                //var excelColumnIndex = new ExcelCellRange(item.ExcelAddress.ToString()).Start.Col;
                var excelColumnIndex = dictExcelAddressCol[item.ExcelAddress];
                dictExcelColumnIndexToModelPropName_All.Add(excelColumnIndex, null);
                string propName = item.Value.ToString();
                PropertyInfo pInfo = type.GetProperty(propName);

                if (pInfo is null)
                {
                    if (dictExcelColumnIndexToModelPropName_Temp.ContainsKey(excelColumnIndex))
                    {
                        var propNameTemp = dictExcelColumnIndexToModelPropName_Temp[excelColumnIndex];
                        //不做属性的 DisplayExcelColumnName = 当前属性的验证 (因为还没想到这个属性是一定要验证的情况)
                        //if (dictModelPropNameToExcelColumnName.ContainsKey(propNameTemp) && dictModelPropNameToExcelColumnName[propNameTemp] == propName)
                        //{
                        //   pInfoTemp = type.GetProperty(propName);
                        //}

                        var pInfoTemp = type.GetProperty(propNameTemp);
                        if (pInfoTemp != null)
                        {
                            propName = propNameTemp;
                            pInfo = pInfoTemp;
                        }
                    }
                }

                //这个If不能放在合并在上面的if中,不然03.11Test的单元测试会不通过
                if (pInfo != null)
                {
                    dictModelPropNameExistsExcelColumn[propName] = true;
                    dictExcelColumnIndexToModelPropName_All[excelColumnIndex] = propName;
                }
            }

            #endregion

            #region 验证 MatchingModel.eq //args.MatchingModel

            var _matchingModelSuccess = false;  //提供的 Matching 参数[这里写死了MatchingModel.eq] 和算出来MatchingModel 有没有交集(默认没有)

            var dictExcelColumnIndexToExcelColName = colNameList.ToDictionary(item => new ExcelCellPoint(item.ExcelAddress).Col, item => item.Value.ToString());
            var _matchingModel = GetMatchingModel(dictExcelColumnIndexToExcelColName, dictExcelColumnIndexToModelPropName_All,
                dictModelPropNameExistsExcelColumn, out List<string> modelPropNotExistsExcelColumn, out List<string> excelColumnIsNotModelProp);
            var _matchingModelValues = Enum.GetValues(typeof(MatchingModel));
            foreach (MatchingModel matchingModelValue in _matchingModelValues)
            {
                if ((MatchingModel.eq & matchingModelValue) == matchingModelValue &&
                    (_matchingModel & matchingModelValue) == matchingModelValue) //提供的 Matching 参数[这里写死了MatchingModel.eq] 和算出来Matching 有重叠
                {
                    _matchingModelSuccess = true;
                    break;
                }
            }
            if (!_matchingModelSuccess)
            {
                #region 获得 listMatchingModelException

                var dictMatchingModelException = new Dictionary<MatchingModel, MatchingModelException>() { };
                //var colNameToCellInfo = colNameList.ToDictionary(item => item.Value.ToString(), item => item);//当excel列有重复,Model没用Attribute,导致Modle 与Excel不匹配, 此时会报错, Demo 在Sample02_7 , 把Model 的Attribute可以复现

                var colNameToCellInfo = new Dictionary<string, List<ExcelCellInfo>>();

                foreach (var colName in colNameList)
                {
                    var colIndex = dictExcelAddressCol[colName.ExcelAddress];
                    var modelpropName = dictExcelColumnIndexToModelPropName_All[colIndex];
                    if (modelpropName != null)
                    {
                        if (!colNameToCellInfo.ContainsKey(modelpropName))
                        {
                            colNameToCellInfo.Add(modelpropName, new List<ExcelCellInfo> { });
                        }
                        colNameToCellInfo[modelpropName].Add(colName);
                    }
                    else
                    {
                        var excelColVaue = colName.Value.ToString();
                        if (!colNameToCellInfo.ContainsKey(excelColVaue))
                        {
                            colNameToCellInfo.Add(excelColVaue, new List<ExcelCellInfo> { });
                            colNameToCellInfo[excelColVaue].Add(colName);
                        }
                        else
                        {
                            ////暂时不考虑多次提供了不存在的列的情况, 即: 不存在的列只能多一共一次,否则报错
                            //throw new Exception($@"当前Excel多次提供了,根据值:{excelColVaue},在Model中找不对应属性,当前列是:{new ExcelCellRange(colName.ExcelAddress.ToString()).Start.R1C1}");
                            if (!colNameToCellInfo.ContainsKey(excelColVaue))
                            {
                                colNameToCellInfo.Add(excelColVaue, new List<ExcelCellInfo> { });
                            }
                            colNameToCellInfo[excelColVaue].Add(colName);
                        }
                    }
                }

                foreach (MatchingModel matchingModelValue in _matchingModelValues)
                {
                    if (matchingModelValue == MatchingModel.eq)
                    {
                        #region excel的哪些列与Model不相等

                        if ((_matchingModel & MatchingModel.eq) != MatchingModel.eq)
                        {
                            continue;
                        }

                        if (dictMatchingModelException.ContainsKey(matchingModelValue))
                        {
                            continue;//如果已经添加过了
                        }

                        if (excelColumnIsNotModelProp.Count == 0 && modelPropNotExistsExcelColumn.Count == 0)
                        {
                            dictMatchingModelException.Add(MatchingModel.eq, new MatchingModelException()
                            {
                                MatchingModel = MatchingModel.eq,
                                ListExcelCellInfoAndModelType = null
                            });
                        }
                        else
                        {
                            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
                            foreach (var colName in excelColumnIsNotModelProp)
                            {
                                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType()
                                {
                                    ExcelCellInfoList = colNameToCellInfo[colName],
                                    ModelType = type
                                });
                            }

                            dictMatchingModelException.Add(MatchingModel.eq, new MatchingModelException()
                            {
                                MatchingModel = MatchingModel.eq,
                                ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType
                            });
                        }

                        #endregion
                    }
                    else if (matchingModelValue == MatchingModel.gt)
                    {
                        if ((_matchingModel & MatchingModel.gt) != MatchingModel.gt)
                        {
                            continue;
                        }

                        if (dictMatchingModelException.ContainsKey(MatchingModel.gt))
                        {
                            continue;
                        }

                        dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, args.ws));
                    }
                    else if (matchingModelValue == MatchingModel.lt)
                    {
                        if ((_matchingModel & MatchingModel.lt) != MatchingModel.lt)
                        {
                            continue;
                        }

                        if (dictMatchingModelException.ContainsKey(MatchingModel.lt))
                        {
                            continue;
                        }

                        dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, args.ws));
                    }
                    else if (matchingModelValue == MatchingModel.neq)
                    {
                        if ((_matchingModel & MatchingModel.neq) != MatchingModel.neq)
                        {
                            continue;
                        }
                        //neq 会调用 gt+ lt ,所以要排除,即 _matchingModel的值 不能是带neq的标志枚举的值
                        if ((_matchingModel & MatchingModel.gt) == MatchingModel.gt)
                        {
                            continue;
                        }

                        if ((_matchingModel & MatchingModel.lt) == MatchingModel.lt)
                        {
                            continue;
                        }

                        #region excel的哪些列与Model不相等

                        //excel的哪些列 不在Model中定义+ model中定义了,但是excel列中却没有

                        if (!dictMatchingModelException.ContainsKey(MatchingModel.gt))
                        {
                            dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, args.ws));
                        }
                        if (!dictMatchingModelException.ContainsKey(MatchingModel.lt))
                        {
                            dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, args.ws));
                        }

                        #endregion

#if DEBUG
                        if ((_matchingModel & MatchingModel.eq) == MatchingModel.eq)
                        {
                            throw new Exception("断言:这里应该是不会进来的,debug下调试看看,进来是什么情况");
                        }
#endif
                    }
                    else
                    {
                        throw new Exception("不支持的MatchingMode值");
                    }
                }

                #endregion

                StringBuilder sb = new StringBuilder();

                //foreach (var matchingModelException in listMatchingModelException)
                foreach (var matchingModelException in dictMatchingModelException.Values)
                {
                    var errMsg = DealMatchingModelException(matchingModelException);
                    sb.Append(errMsg);
                }
                if (sb.Length > 0)
                {
                    throw new MatchingModelException(sb.ToString());
                }

                throw new Exception("验证未通过,程序有bug");
            }

            #endregion

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList is null
                ? GetExcelListArgs.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;

            #region 初始化内置Attribute 和 检查模型属性

            var dictUnique = new Dictionary<string, List<string>>();//属性的 UniqueAttribute 的内置实现
            var dictCustomerModelType = new Dictionary<PropertyInfo, ICustomersModelType>();

            foreach (var excelCellInfo in colNameList)
            {
                string propName = ExcelAddressHelper.GetPropName(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All);
                if (string.IsNullOrEmpty(propName))
                {
                    continue;
                }

                var pInfo = ReflectionHelper.GetPropertyInfo(propName, type);

                #region 初始化Attr要处理相关的数据

                var uniqueAttrs = ReflectionHelper.GetAttributeForProperty<UniqueAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (uniqueAttrs.Length > 0)
                {
                    dictUnique.Add(pInfo.Name, new List<string>());
                }

                var isInList = ICustomersModelTypeList.Find(a => a == pInfo) != null;
                if (isInList)
                {
                    ICustomersModelType p = (ICustomersModelType)pInfo.GetValue(args.Model);

                    dictCustomerModelType.Add(pInfo, p);
                }
                #endregion
            }

            #endregion

            #region 获得 list

            ExcelCellInfoNeedTo(args.ReadCellValueOption, out var toTrim, out var toMergeLine, out var toDBC);

            var allRowExceptions = args.GetList_NeedAllException ? new List<Exception>() : null;

            Func<object[], object> deletgateCreateInstance = ExpressionTreeHelper.BuildDeletgateCreateInstance(type, new Type[0]);

            var dynamicCalcStep = DynamicCalcStep(args.ScanLine);
            int row = args.DataRowStart;

            while (true)//遍历行, 异常或者出现空行,触发break;
            {
                if (args.ws.Row(row).Hidden)
                {
                    throw new Exception($@"工作簿:'{args.ws.Name}'不允许存在隐藏行,检测到第{row}行是隐藏行");
                }

                //判断整行数据是否都没有数据
                bool isNoDataAllColumn = true;

                //Sample02_3,12000的数据
                //T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转  1.2-2.1秒
                //T model = type.CreateInstance<T>();//3秒+
                T model = (T)deletgateCreateInstance(null); //上面的方法给拆开来 . 1.1-1.4
                var thisRowExceptions = new List<Exception>();
                foreach (var excelCellInfo in colNameList)
                {
                    string propName = ExcelAddressHelper.GetPropName(excelCellInfo.ExcelAddress, dictExcelAddressCol,
                        dictExcelColumnIndexToModelPropName_All);
                    if (string.IsNullOrEmpty(propName))
                    {
                        continue;
                    }

                    var pInfo = ReflectionHelper.GetPropertyInfo(propName, type);
                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                    var value = GetValue(pInfo, row, col);

                    #region 全局处理值,如特性

                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);
                    if (!valueIsNullOrEmpty)
                    {
                        if (isNoDataAllColumn)
                        {
                            isNoDataAllColumn = false;
                        }

                        #region 判断每个单元格的开头

                        if (args.EveryCellPrefix?.Length > 0)
                        {
                            var indexof = value.IndexOf(args.EveryCellPrefix, StringComparison.Ordinal);
                            if (indexof == -1)
                            {
                                throw new ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
                            }
                            value = value.RemovePrefix(args.EveryCellPrefix);
                        }
                        #endregion

                        #region 对每个单元格进行值的替换

                        if (everyCellReplace != null)
                        {
                            foreach (var replaceItem in everyCellReplace)
                            {
                                if (!value.Contains(replaceItem.Key))
                                {
                                    continue;
                                }
                                var cellReplaceOldValue = replaceItem.Key;
                                var cellReplaceNewValue = replaceItem.Value ?? "";
                                if (cellReplaceOldValue?.Length > 0)
                                {
                                    value = value.Replace(cellReplaceOldValue, cellReplaceNewValue);
                                }
                            }
                        }
                        #endregion

                        #region 对每个单元格进行处理

                        if (toTrim)
                        {
                            value = value.Trim();
                        }
                        if (toMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (toDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion

                        #region 处理内置的Attribute

                        var propAttrs = pInfo.GetCustomAttributes();

                        foreach (var item in propAttrs)
                        {
                            var attrType = item.GetType();

                            if (attrType == typeof(UniqueAttribute))
                            {
                                var uniqueAttrInfo = item;
                                if (uniqueAttrInfo != null)
                                {
                                    if (dictUnique[pInfo.Name].Contains(value))
                                    {
                                        string exception_msg = string.IsNullOrEmpty(((UniqueAttribute)uniqueAttrInfo).ErrorMessage)
                                            ? $@"属性'{pInfo.Name}'的值:'{value}'出现了重复"
                                            : ((UniqueAttribute)uniqueAttrInfo).ErrorMessage;
                                        throw new ArgumentException(exception_msg, pInfo.Name);
                                    }
                                    dictUnique[pInfo.Name].Add(value);
                                }
                            }
                            if (dictCustomerModelType.ContainsKey(pInfo))
                            {
                                var customerModelType = dictCustomerModelType[pInfo];
                                if (customerModelType.HasAttribute)
                                {
                                    customerModelType.RunAttribute(item, pInfo, model, value);
                                }
                            }
                        }

                        #endregion
                    }
                    #endregion

                    Exception exception = null;
                    try
                    {
                        //验证继承 System.ComponentModel.DataAnnotations 的那些特性们
                        if (pInfo.IsDefined(typeof(ValidationAttribute), true))
                        {
                            object[] validAttrs = pInfo.GetCustomAttributes(typeof(ValidationAttribute), true);
                            foreach (ValidationAttribute validAttr in validAttrs)
                            {
                                validAttr.Validate(value, name: null);
                            }
                        }

                        if (ICustomersModelTypeList.Contains(pInfo))
                        {
                            dictCustomerModelType[pInfo].SetModelValue(pInfo, model, value);
                        }
                        else
                        {
                            GetList_SetModelValue(pInfo, model, value);
                        }
                    }
                    catch (ArgumentException e)
                    {
                        exception = new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}", e);
                    }
                    catch (ValidationException e)
                    {
                        exception = new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}({pInfo.Name}:{e.Message})", e);
                    }
                    catch (Exception e)
                    {
                        exception = e;
                    }
                    finally
                    {
                        if (exception != null)
                        {
                            thisRowExceptions.Add(exception);
                        }
                    }
                }

                //1.添加Step,准备读取下一行数据
                if (dynamicCalcStep)
                {
                    //while里面动态计算
                    if (ExcelWorksheetHelper.IsMergeCell(args.ws, row, col: args.DataColStart, out var mergeCellAddress))
                    {
                        row += new ExcelAddress(mergeCellAddress).Rows;//按第一列合并的行数进行step的增加
                    }
                    else
                    {
                        row += 1;
                    }
                }
                else
                {
                    row += 1;
                }

                if (isNoDataAllColumn)
                {
                    break; //出现空行,无视异常,结束while(true)循环,模版读取结束
                }

                //2.如果有异常,抛出异常,不对下一行进行读取
                if (thisRowExceptions.Count > 0)
                {
                    if (args.GetList_NeedAllException)
                    {
                        allRowExceptions.AddRange(thisRowExceptions);
                    }
                    else
                    {
                        throw thisRowExceptions[0];
                    }
                }

                if (args.WhereFilter is null || args.WhereFilter.Invoke(model))
                {
                    if (args.HavingFilter is null || args.HavingFilter.Invoke(model))
                    {
                        yield return model;
                    }
                }
            }

            var keyWithExceptionMessageStart = "无效的单元格:";
            if (allRowExceptions != null && allRowExceptions.Count > 0)
            {
                bool allExceptionIsArgumentException = true;
                var errGroupMsg = new Dictionary<string, List<string>>();

                foreach (var ex in allRowExceptions)
                {
                    if (!(ex is ArgumentException))
                    {
                        allExceptionIsArgumentException = false;
                        break;
                    }
                    if (!((ArgumentException)ex).Message.StartsWith(keyWithExceptionMessageStart))
                    {
                        allExceptionIsArgumentException = false;
                        break;
                    }

                    var excelCellAddress = ((ArgumentException)ex).Message.RemovePrefix(keyWithExceptionMessageStart);
                    var exceptionMessage = ((ArgumentException)ex).InnerException.Message;
                    if (!errGroupMsg.ContainsKey(exceptionMessage))
                    {
                        errGroupMsg.Add(exceptionMessage, new List<string>());
                    }

                    errGroupMsg[exceptionMessage].Add(excelCellAddress);
                }

                if (!allExceptionIsArgumentException)
                {
                    throw new AggregateException(allRowExceptions);
                }

                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                foreach (KeyValuePair<string, List<string>> msg in errGroupMsg)
                {
                    sb.Append(msg.Key);
                    sb2.Clear();
                    if (args.GetList_ErrorMessage_OnlyShowColomn)
                    {
                        var cols = new List<string>();
                        foreach (string excelCellAddress in msg.Value)
                        {
                            cols.Add(ExcelCellPoint.R1C1FormulasReverse(new ExcelCellAddress(excelCellAddress).Column));
                        }

                        foreach (var col in cols.Distinct())
                        {
                            sb2.Append(col).Append("列,");
                        }
                    }
                    else
                    {
                        foreach (var excelCellAddress in msg.Value)
                        {
                            sb2.Append(excelCellAddress).Append(",");
                        }
                    }
                    sb.AppendLine($"({sb2.RemoveLastChar(',')}),");
                }
                //var argumentExceptionMsg = sb.RemoveLastChar(Environment.NewLine).RemoveLastChar(',').AppendLine().ToString();
                var argumentExceptionMsg = sb.Replace(",", "", sb.Length - Environment.NewLine.Length - 1, 1).ToString();//与上面等价

                throw new ArgumentException(argumentExceptionMsg);
            }

            #endregion
        }

        /// <summary>
        /// 是否是动态计算步长
        /// </summary>
        /// <param name="scanLine"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static bool DynamicCalcStep(ScanLine scanLine)
        {
            if (scanLine == ScanLine.SingleLine)
            {
                return false;
            }

            if (scanLine == ScanLine.MergeLine)
            {
                return true; //在代码的while中进行动态计算
            }

            throw new Exception("不支持的ScanLine");
        }

        /// <summary>
        /// 从Excel 中获得符合C# 类属性定义的列名集合,内部会修改DataColEnd的值
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static List<ExcelCellInfo> GetExcelColumnOfModel(GetExcelListArgs args)
        {
            List<string> colNameList = null;
            Dictionary<string, int> nameRepeatCounter = null;
            if (args.POCO_Property_AutoRename_WhenRepeat)
            {
                colNameList = new List<string>();
                nameRepeatCounter = new Dictionary<string, int>();
            }

            var list = new List<ExcelCellInfo>();
            int col = args.DataColStart;
            int dataColEndActual = 0;
            while (col <= args.DataColEnd)
            {
                if (args.ws.Column(col).Hidden)
                {
                    throw new Exception($@"工作簿:'{args.ws.Name}'不允许存在隐藏列,检测到第{ExcelCellPoint.R1C1FormulasReverse(col)}列是隐藏列");
                }
                ExcelAddress ea;
                int newDataColEndActual;
                var isMergeCell = ExcelWorksheetHelper.IsMergeCell(args.ws, args.DataTitleRow, col, out var mergeCellAddress);
                if (isMergeCell)
                {
                    ea = new ExcelAddress(mergeCellAddress);
                    newDataColEndActual = col + new ExcelCellRange(mergeCellAddress).IntervalCol;
                    col += ea.Columns;
                }
                else
                {
                    ea = new ExcelAddress(args.DataTitleRow, col, args.DataTitleRow, col);
                    newDataColEndActual = col;
                    col += 1;
                }

                var colName = ExcelWorksheetHelper.GetMergeCellText(args.ws, ea.Start.Row, ea.Start.Column);
                if (string.IsNullOrEmpty(colName))
                {
                    break;
                }

                colName = NamingHelper.ExtractName(colName);
                if (string.IsNullOrEmpty(colName))
                {
                    break;
                }

                dataColEndActual = newDataColEndActual;

                if (args.POCO_Property_AutoRename_WhenRepeat)
                {
                    NamingHelper.AutoRename(colNameList, nameRepeatCounter, colName, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
                }
                list.Add(new ExcelCellInfo()
                {
                    WorkSheet = args.ws,
                    ExcelAddress = ea,
                    Value = colName,
                });
            }
            if (args.DataColEnd == EPPlusConfig.MaxCol07)//当前是恒成立,因为DataColEnd 是internal
            {
                args.DataColEnd = dataColEndActual;
            }
            else
            {
                if (args.DataColEnd != dataColEndActual) //当前 DataColEnd 是internal 的,不会执行到这里,这里是防止以后程序修改而写的.
                {
                    throw new Exception("非预期的值,请检查当前程序或使用代码.");
                }
            }

            if (args.POCO_Property_AutoRename_WhenRepeat)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    list[i].Value = colNameList[i];
                }
            }

            if (list.Count == 0)
            {
                throw new Exception("未读取到单元格标题");
            }

            return list;
        }

        private static void GetList_SetModelValue<T>(PropertyInfo pInfo, T model, string value) where T : class, new()
        {
            var pInfo_PropertyType = pInfo.PropertyType;
            #region string

            if (pInfo_PropertyType == typeof(string))
            {
                pInfo.SetValue(model, value);
                //pInfo.SetValue(model, ws.Cells[row, col].Text);
                //type.GetProperty(colName)?.SetValue(model, ws.Cells[row, col].Text);
                return;
            }
            #endregion
            #region Boolean

            var isNullable_Boolean = pInfo_PropertyType == typeof(Boolean?);

            if (isNullable_Boolean && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }

            if (isNullable_Boolean || pInfo_PropertyType == typeof(Boolean))
            {
                if (!Boolean.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的Boolean值", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 Boolean。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region DateTime

            var isNullable_DateTime = pInfo_PropertyType == typeof(DateTime?);
            if (isNullable_DateTime && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }

            if (isNullable_DateTime || pInfo_PropertyType == typeof(DateTime))
            {
                if (!DateTime.TryParse(value, out var result))
                {
                    if (!double.TryParse(value, out var resultDouble))
                    {
                        throw new ArgumentException("无效的日期", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 DateTime。"));
                    }
                    //excel日期用数字保存的

                    //在百度看到
                    //excel与VBA开始点有差别:
                    //excel开始点: 1900-1-1 序号为1
                    //vba开始点:1899-12-31 序号为1
                    //原因是excel把1900-2月错误地当29天处理,所以VBA后来自己修改了这个错误,以能与excel相适应.从1900年3月1日开始,VBA与Excel的序号才开始一致.

                    //数字转日期: //参考文章 : https://docs.microsoft.com/zh-cn/dotnet/api/system.datetime.fromoadate   该方法测试发现 DateTime.FromOADate(d)  d值必须>= -657434.999999999941792(后面还能添加数字,未测试) && d<=2958465.999999994(后面还能添加数字,没测试)
                    //但是在excel 日期最多精确到毫秒3位, 即 yyyy-MM-dd HH:mm:ss.000,对应的日期值的范围是 [1,2958465.99999999],且不能包含[60,61)
                    //Excel数值对应的日期
                    //0                 对应 1900-01-00 00:00:00.000   (日期不对)
                    //1                 对应 1900-01-01 00:00:00.000
                    //60                对应 1900-02-29 00:00:00.000   (日期不对)
                    //2958465.99999999  对应 9999-12-31 23:59:59.999  但是  DateTime.Parse("9999-12-31 23:59:59.999").ToOADate()  2958465.9999999884
                    if (resultDouble >= 1 && resultDouble < 60)
                    {
                        result = DateTime.FromOADate(resultDouble + 1);
                    }
                    else if (resultDouble >= 61 && resultDouble <= 2958465.99999999)
                    {
                        result = DateTime.FromOADate(resultDouble);
                    }
                    else
                    {
                        throw new ArgumentException("无效的日期", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 DateTime。{value}必须是在[1,60) 或 [61,2958465.99999999]之间的值"));
                    }
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region sbyte

            var isNullable_sbyte = pInfo_PropertyType == typeof(sbyte?);
            if (isNullable_sbyte && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_sbyte || pInfo_PropertyType == typeof(sbyte))
            {
                if (!sbyte.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 sbyte。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region byte

            var isNullable_byte = pInfo_PropertyType == typeof(byte?);
            if (isNullable_byte && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_byte || pInfo_PropertyType == typeof(byte))
            {
                if (!byte.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 byte。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region UInt16

            var isNullable_UInt16 = pInfo_PropertyType == typeof(UInt16?);
            if (isNullable_UInt16 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_UInt16 || pInfo_PropertyType == typeof(UInt16))
            {
                if (!UInt16.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 UInt16。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region Int16

            var isNullable_Int16 = pInfo_PropertyType == typeof(Int16?);
            if (isNullable_Int16 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_Int16 || pInfo_PropertyType == typeof(Int16))
            {
                if (!Int16.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 Int16。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region UInt32

            var isNullable_UInt32 = pInfo_PropertyType == typeof(UInt32?);
            if (isNullable_UInt32 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_UInt32 || pInfo_PropertyType == typeof(UInt32))
            {
                if (!UInt16.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 UInt32。"));
                }
                pInfo.SetValue(model, result);
                return;
            }

            #endregion
            #region Int32

            var isNullable_Int32 = pInfo_PropertyType == typeof(Int32?);
            if (isNullable_Int32 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_Int32 || pInfo_PropertyType == typeof(Int32))
            {
                if (!Int32.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 Int32。"));
                }
                pInfo.SetValue(model, result);
                return;
            }

            #endregion
            #region UInt64

            var isNullable_UInt64 = pInfo_PropertyType == typeof(UInt64?);
            if (isNullable_UInt64 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_UInt64 || pInfo_PropertyType == typeof(UInt64))
            {
                if (!UInt64.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 UInt64。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region Int64

            var isNullable_Int64 = pInfo_PropertyType == typeof(Int64?);
            if (isNullable_Int64 && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_Int64 || pInfo_PropertyType == typeof(Int64))
            {
                if (!Int64.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 Int64。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region float

            var isNullable_float = pInfo_PropertyType == typeof(float?);
            if (isNullable_float && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_float || pInfo_PropertyType == typeof(float))
            {
                if (!float.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 float。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region double

            var isNullable_double = pInfo_PropertyType == typeof(double?);
            if (isNullable_double && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_double || pInfo_PropertyType == typeof(double))
            {
                if (!double.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 double。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region decimal

            var isNullable_decimal = pInfo_PropertyType == typeof(decimal?);
            if (isNullable_decimal && (value is null || value.Length <= 0))
            {
                pInfo.SetValue(model, null);
                return;
            }
            if (isNullable_decimal || pInfo_PropertyType == typeof(decimal))
            {
                if (!Decimal.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 decimal。"));
                }
                pInfo.SetValue(model, result);
                return;
            }
            #endregion
            #region Enum

            bool isNullable_Enum = Nullable.GetUnderlyingType(pInfo_PropertyType)?.IsEnum == true;
            if (isNullable_Enum)
            {
                if (value is null || value.Length <= 0)
                {
                    pInfo.SetValue(model, null);
                    return;
                }
                value = NamingHelper.ExtractName(value);
                var enumType = pInfo_PropertyType.GetProperty("Value").PropertyType;
                TryThrowExceptionForEnum(pInfo, model, value, enumType, pInfo_PropertyType);
                var enumValue = Enum.Parse(enumType, value);
                pInfo.SetValue(model, enumValue);
                return;
            }
            if (pInfo_PropertyType.IsEnum)
            {
                if ((value is null || value.Length <= 0))
                {
                    throw new ArgumentException($@"无效的{pInfo_PropertyType.FullName}枚举值", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 {pInfo_PropertyType}(Enum类型)"));
                }
                value = NamingHelper.ExtractName(value);
                TryThrowExceptionForEnum(pInfo, model, value, pInfo_PropertyType, pInfo_PropertyType);
                var enumValue = Enum.Parse(pInfo_PropertyType, value);
                pInfo.SetValue(model, enumValue);
                return;
            }
            #endregion

            throw new Exception("GetList_SetModelValue()时遇到未处理的类型!!!请完善程序");
        }

        private static void TryThrowExceptionForEnum<T>(PropertyInfo pInfo, T model, string value, Type enumType, Type pInfoType) where T : class, new()
        {
            var isDefined = Enum.IsDefined(enumType, value);
            if (isDefined)
            {
                return;
            }
            var attrs = ReflectionHelper.GetAttributeForProperty<EnumUndefinedAttribute>(pInfo.DeclaringType, pInfo.Name);
            if (attrs.Length == 1)
            {
                //使用自定义消息
                var attr = (EnumUndefinedAttribute)attrs[0];
                if (attr.Args is null || attr.Args.Length <= 0)
                {
                    if (string.IsNullOrEmpty(attr.ErrorMessage))
                    {
                        throw new ArgumentException($"Value值:'{value}'在枚举值:'{pInfoType.FullName}'中未定义,请检查!!!");
                    }

                    throw new ArgumentException(attr.ErrorMessage);
                }

                var message = FormatAttributeMsg(pInfo.Name, model, value, attr.ErrorMessage, attr.Args);
                throw new ArgumentException(message);
            }
            else
            {
                //使用枚举类型内置的消息: 未找到请求的值“xxx”。
                //throw new ArgumentException($"Value值:'{value}'在枚举值:'{pInfoType.FullName}'中未定义,请检查!!!");
            }
        }

        /// <summary>
        /// 格式化Attribute的错误消息
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="pInfo_Name"></param>
        /// <param name="model"></param>
        /// <param name="value"></param>
        /// <param name="attrErrorMessage"></param>
        /// <param name="attrArgs"></param>
        /// <returns></returns>
        public static string FormatAttributeMsg<T>(string pInfo_Name, T model, string value, string attrErrorMessage, string[] attrArgs) where T : class, new()
        {
            //拼接ErrorMessage
            var allProp = ReflectionHelper.GetProperties<T>();

            for (int i = 0; i < attrArgs.Length; i++)
            {
                var propertyName = attrArgs[i];
                if (string.IsNullOrEmpty(propertyName))
                {
                    continue;
                }

                //注:如果占位符这是常量且刚好和属性名一致,请把占位符拆成多个占位符使用
                if (propertyName == pInfo_Name)
                {
                    attrArgs[i] = value;
                }
                else
                {
                    var prop = ReflectionHelper.GetProperty(allProp, propertyName, true);
                    if (prop is null)
                    {
                        continue;
                    }

                    attrArgs[i] = prop.GetValue(model).ToString();
                }
            }

            string message = string.Format(attrErrorMessage, attrArgs);
            return message;
        }

        private static void ExcelCellInfoNeedTo(ReadCellValueOption readCellValueOption, out bool toTrim, out bool toMergeLine,
            out bool toDBC)
        {
            toTrim = (readCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            toMergeLine = (readCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            toDBC = (readCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;
        }


        /// <summary>
        /// 读取excel, 返回一个DataTable
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static DataTable GetDataTable(GetExcelListArgs<DataRow> args)
        {
            ExcelWorksheet ws = args.ws;
            int rowStart = args.DataRowStart;
            if (rowStart <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowStart}'错误,值应该大于0");
            }

            if (args.DataTitleRow <= 0)
            {
                throw new ArgumentException($@"数据起始行的标题行值'{args.DataTitleRow}'错误,值应该大于0");
            }

            var colNameList = GetExcelColumnOfModel(args);
            if (colNameList.Count == 0)
            {
                throw new Exception("未读取到单元格标题");
            }
            var dictExcelAddressCol = colNameList.ToDictionary(item => item.ExcelAddress, item => new ExcelCellPoint(item.ExcelAddress).Col);

            #region 获得字典

            DataTable dt = new DataTable();

            foreach (var item in colNameList)
            {
                dt.Columns.Add(item.Value.ToString());
            }

            #endregion

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList is null
                ? GetExcelListArgs.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;

            #region 获得 list

            int row = rowStart;
            bool dynamicCalcStep = DynamicCalcStep(args.ScanLine);
            ExcelCellInfoNeedTo(args.ReadCellValueOption, out var toTrim, out var toMergeLine, out var toDBC);
            while (true)
            {
                bool isNoDataAllColumn = true;//判断整行数据是否都没有数据
                var dr = dt.NewRow();

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    string propName = excelCellInfo.Value.ToString();

                    if (string.IsNullOrEmpty(propName))
                    {
                        continue;//理论上,这种情况不存在,即使存在了,也要跳过
                    }

                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

                    string value = ExcelWorksheetHelper.GetMergeCellText(ws, row, col);
                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);

                    if (!valueIsNullOrEmpty)
                    {
                        isNoDataAllColumn = false;

                        #region 判断每个单元格的开头

                        if (args.EveryCellPrefix?.Length > 0)
                        {
                            var indexof = value.IndexOf(args.EveryCellPrefix, StringComparison.Ordinal);
                            if (indexof == -1)
                            {
                                throw new ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
                            }
                            value = value.RemovePrefix(args.EveryCellPrefix);
                        }
                        #endregion

                        #region 对每个单元格进行值的替换

                        if (everyCellReplace != null)
                        {
                            foreach (var item in everyCellReplace)
                            {
                                if (!value.Contains(item.Key))
                                {
                                    continue;
                                }
                                var everyCellReplaceOldValue = item.Key;
                                var everyCellReplaceNewValue = item.Value ?? "";
                                if (everyCellReplaceOldValue?.Length > 0)
                                {
                                    value = value.Replace(everyCellReplaceOldValue, everyCellReplaceNewValue);
                                }
                            }
                        }
                        #endregion

                        #region 对每个单元格进行处理

                        if (toTrim)
                        {
                            value = value.Trim();
                        }
                        if (toMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (toDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion
                    }

                    dr[propName] = value;//赋值
                }

                if (isNoDataAllColumn)
                {
                    if (row == rowStart)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    break; //出现空行,读取模版结束
                }
                //else
                if (args.WhereFilter is null || args.WhereFilter.Invoke(dr))
                {
                    dt.Rows.Add(dr);
                }
                if (dynamicCalcStep)
                {
                    if (ExcelWorksheetHelper.IsMergeCell(ws, row, col: 1, out var mergeCellAddress))
                    {
                        row += new ExcelAddress(mergeCellAddress).Rows;
                    }
                    else
                    {
                        row += 1;
                    }
                }
                else
                {
                    row += 1;
                }
            }

            #endregion

            var result = args.HavingFilter is null
                 ? dt
                 : dt.AsEnumerable()
                     .Where(item => args.HavingFilter.Invoke(item))
                     .CopyToDataTable();

            return result;
        }

        private static string DealMatchingModelException(MatchingModelException matchingModelException)
        {
            if ((matchingModelException.MatchingModel & MatchingModel.eq) == MatchingModel.eq)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType is null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有多提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版提供了多余的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"{excelCellInfo.ExcelAddress}({excelCellInfo.Value}),");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();
            }
            if ((matchingModelException.MatchingModel & MatchingModel.gt) == MatchingModel.gt)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType is null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有少提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版多提供了model属性中不存在的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"{excelCellInfo.ExcelAddress}({excelCellInfo.Value}),");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();
            }
            if ((matchingModelException.MatchingModel & MatchingModel.lt) == MatchingModel.lt)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType is null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有多提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版少提供了model属性中定义的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"'{excelCellInfo.Value}',");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();
            }
            throw new Exception($@"参数{nameof(matchingModelException)},不支持的MatchingMode值");
        }

        /// <summary>
        ///  model的哪些属性是在excel中没有定义的 + model中没有定义
        /// </summary>
        /// <param name="excelColumnIsNotModelProp"></param>
        /// <param name="type"></param>
        /// <param name="colNameToCellInfo"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static MatchingModelException GetMatchingModelExceptionCase_lt(List<string> excelColumnIsNotModelProp,
            Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
        {
            if (excelColumnIsNotModelProp.Count <= 0)
            {
                return new MatchingModelException { MatchingModel = MatchingModel.lt, ListExcelCellInfoAndModelType = null };
            }

            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
            foreach (var propName in excelColumnIsNotModelProp)
            {
                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType
                {
                    ModelType = type,
                    ExcelCellInfoList = colNameToCellInfo.ContainsKey(propName)
                        ? colNameToCellInfo[propName]
                        : new List<ExcelCellInfo> { new ExcelCellInfo { Value = propName, ExcelAddress = null, WorkSheet = ws } }
                });
            }

            return new MatchingModelException() { MatchingModel = MatchingModel.lt, ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType };
        }

        /// <summary>
        /// excel的哪些列是在Model中定义了却没有(即,model中缺少的列) + model中没有定义
        /// </summary>
        /// <param name="modelPropNotExistsExcelColumn"></param>
        /// <param name="type"></param>
        /// <param name="colNameToCellInfo"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static MatchingModelException GetMatchingModelExceptionCase_gt(List<string> modelPropNotExistsExcelColumn, Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
        {
            if (modelPropNotExistsExcelColumn.Count <= 0)
            {
                return new MatchingModelException { MatchingModel = MatchingModel.eq, ListExcelCellInfoAndModelType = null };
            }

            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
            foreach (var colName in modelPropNotExistsExcelColumn)
            {
                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType
                {
                    ModelType = type,
                    ExcelCellInfoList = colNameToCellInfo.ContainsKey(colName)
                        ? colNameToCellInfo[colName]
                        : new List<ExcelCellInfo> { new ExcelCellInfo { Value = colName, ExcelAddress = null, WorkSheet = ws } }
                });
            }

            return new MatchingModelException { MatchingModel = MatchingModel.gt, ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType };
        }

        private static MatchingModel GetMatchingModel(
            Dictionary<int, string> dictExcelColumnIndexToExcelColName,
            Dictionary<int, string> dictExcelColumnIndexToModelPropNameAll,
            Dictionary<string, bool> dictModelPropNameExistsExcelColumn,
            out List<string> modelPropNotExistsExcelColumn, out List<string> excelColumnIsNotModelProp)
        {
            if (dictExcelColumnIndexToModelPropNameAll is null)
            {
                throw new ArgumentNullException(nameof(dictExcelColumnIndexToModelPropNameAll));
            }
            if (dictModelPropNameExistsExcelColumn is null)
            {
                throw new ArgumentNullException(nameof(dictModelPropNameExistsExcelColumn));
            }

            modelPropNotExistsExcelColumn = new List<string>();//model属性不在excel列中
            excelColumnIsNotModelProp = new List<string>();//excel列不是model属性

            if (dictExcelColumnIndexToModelPropNameAll.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.eq;
            }

            if (dictExcelColumnIndexToModelPropNameAll.Keys.Count > 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.neq | MatchingModel.gt;
            }

            if (dictExcelColumnIndexToModelPropNameAll.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count > 0)
            {
                return MatchingModel.neq | MatchingModel.lt;
            }

            foreach (var excelColumnIndex in dictExcelColumnIndexToModelPropNameAll.Keys)
            {
                if (dictExcelColumnIndexToModelPropNameAll[excelColumnIndex] is null)
                {
                    modelPropNotExistsExcelColumn.Add(dictExcelColumnIndexToExcelColName[excelColumnIndex]);
                }
            }

            foreach (var modelPropName in dictModelPropNameExistsExcelColumn.Keys)
            {
                if (!dictModelPropNameExistsExcelColumn[modelPropName])
                {
                    excelColumnIsNotModelProp.Add(modelPropName);
                }
            }

            //这里要出重,因为该方法外层的 colNameToCellInfo 对象的类型从 Dictionary<string, ExcelCellInfo>() 改为了 Dictionary<string, List<ExcelCellInfo>>
            if (modelPropNotExistsExcelColumn.Count > 0)
            {
                modelPropNotExistsExcelColumn = modelPropNotExistsExcelColumn.Distinct().ToList();
            }
            if (excelColumnIsNotModelProp.Count > 0)
            {
                excelColumnIsNotModelProp = excelColumnIsNotModelProp.Distinct().ToList();
            }

            if (excelColumnIsNotModelProp.Count == 0 && modelPropNotExistsExcelColumn.Count == 0)
            {
                return MatchingModel.eq;
            }

            if (excelColumnIsNotModelProp.Count > 0 && modelPropNotExistsExcelColumn.Count > 0)
            {
                return MatchingModel.neq;
            }

            if (modelPropNotExistsExcelColumn.Count > 0)//model属性多,即, excel列的数量 比model属性数量少
            {
                return MatchingModel.neq | MatchingModel.gt;
            }

            if (excelColumnIsNotModelProp.Count > 0)
            {
                return MatchingModel.neq | MatchingModel.lt;
            }

            throw new Exception(nameof(GetMatchingModel) + "程序不对,需要修改");
        }

        #endregion

        #region 设置单元格值

        public static void SetWorksheetCellValue(ExcelRange cell, string cellValue)
        {
            ExcelRangeHelper.SetWorksheetCellValue(cell, cellValue);
        }

        #endregion

        #region 读取excel配置

        /// <summary>
        /// 从Excel中获取默认的配置信息
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetIndex">第几个workSheet作为模版,从1开始</param>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, int workSheetIndex)
        {
            if (workSheetIndex <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            }
            var worksheet = GetExcelWorksheet(excelPackage, workSheetIndex);
            EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);
            SetConfigBodyFromExcel_OtherPara(config, worksheet);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetName"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, string workSheetName)
        {
            if (workSheetName is null)
            {
                throw new ArgumentNullException(nameof(workSheetName));
            }
            var worksheet = GetExcelWorksheet(excelPackage, workSheetName);
            EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);
            EPPlusHelper.SetConfigBodyFromExcel_OtherPara(config, worksheet);
        }

        /// <summary>
        ///  这个不能在FillData里面算, 会有问题
        /// </summary>
        /// <param name="config"></param>
        /// <param name="worksheet"></param>
        private static void SetConfigBodyFromExcel_OtherPara(EPPlusConfig config, ExcelWorksheet worksheet)
        {
            foreach (var configItem in config.Body.ConfigList)
            {
                var nth = configItem.Nth;

                var allConfigInterval = config.Body[nth].Option.ConfigLine.Count; //配置共计用了多少列, 默认: 1个配置用了1列

                var mergedCellsList = worksheet.MergedCells.Where(a => a != null).ToList();
                foreach (var configCellInfo in config.Body[nth].Option.ConfigLine)
                {
                    if (worksheet.Cells[configCellInfo.Address].Merge) //item.Address  D4
                    {
                        var addressPrecise = ExcelWorksheetHelper.GetMergeCellAddressPrecise(worksheet, configCellInfo.Address); //D4:E4格式的
                        allConfigInterval += new ExcelCellRange(addressPrecise).IntervalCol;

                        configCellInfo.FullAddress = addressPrecise;
                        configCellInfo.IsMergeCell = true;
                    }
                    else
                    {
                        var mergeCellAddress = mergedCellsList.FirstOrDefault(a => a.Contains(configCellInfo.Address));
                        if (mergeCellAddress != null)
                        {
                            allConfigInterval += new ExcelCellRange(mergeCellAddress).IntervalCol;
                            configCellInfo.FullAddress = mergeCellAddress;
                            configCellInfo.IsMergeCell = true;
                        }
                        else
                        {
                            configCellInfo.FullAddress = configCellInfo.Address;
                            configCellInfo.IsMergeCell = false;
                        }
                    }
                }

                config.Body[nth].Option.ConfigLineInterval = allConfigInterval;
            }
        }

        /// <summary>
        /// 让 sheet.Cells.Value 强制从A1单元格开始
        /// </summary>
        /// <param name="sheet"></param>
        public static void SetSheetCellsValueFromA1(ExcelWorksheet sheet)
        {
            //这个可以解决这个问题:
            //描述如下:创建一个excel,在C7,C8,C9,10单元格写入一些字符串,
            //然后通过 sheet.Cells.Value 我获得是object[4,3]的数组,
            //但我要的是object[10,3]的数组
            //解决方法是设置单元的A1
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value is null)
            {
                cellA1.Value = null;
            }
        }

        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static ExcelCellPoint GetFirstValueCellPoint(ExcelWorksheet ws) => (ExcelCellPoint)GetFirstValueCellInfo<ExcelCellPoint>(ws);

        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static ExcelCellRange GetFirstValueCellRange(ExcelWorksheet ws) => (ExcelCellRange)GetFirstValueCellInfo<ExcelCellRange>(ws);

        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <typeparam name="TOut"></typeparam>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static object GetFirstValueCellInfo<TOut>(ExcelWorksheet ws)
        {
            EPPlusHelper.SetSheetCellsValueFromA1(ws);
            object[,] arr = ws.Cells.Value as object[,];
            if (arr is null)
            {
                throw new ArgumentNullException(nameof(arr));
            }

            var returnType = typeof(TOut);
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] is null)
                    {
                        continue;
                    }

                    if (arr[i, j].ToString().Length <= 0)
                    {
                        continue;
                    }

                    if (returnType == typeof(ExcelCellPoint))
                    {
                        var cell = new ExcelCellPoint(i + 1, j + 1);
                        return cell;
                    }
                    if (returnType == typeof(ExcelCellRange))
                    {
                        var mergeCellAddress = ExcelWorksheetHelper.GetMergeCellAddressPrecise(ws, i + 1, j + 1);
                        var cell = new ExcelCellRange(mergeCellAddress);
                        return cell;
                    }

                    throw new ArgumentOutOfRangeException(nameof(returnType), $@"不支持的参数{nameof(returnType)}类型:{returnType}");
                }
            }

            if (returnType == typeof(ExcelCellPoint))
            {
                return new ExcelCellPoint();
            }

            if (returnType == typeof(ExcelCellRange))
            {
                return new ExcelCellPoint();
            }

            throw new ArgumentOutOfRangeException(nameof(returnType), $@"不支持的参数{nameof(returnType)}类型:{returnType}");
        }

        /// <summary>
        /// 设置默认的配置
        /// </summary>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetDefaultConfigFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            SetSheetCellsValueFromA1(sheet);
            config.Head = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$th") };
            SetConfigBodyFromExcel(config, sheet);
            config.Foot = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$tf") };
        }

        /// <summary>
        /// 设置sheetBody配置
        /// </summary>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetConfigBodyFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");

            var sheetMergedCellsList = sheet.MergedCells.ToList();

            var bodyConfigCache = new Dictionary<int, EPPlusConfigBodyConfig>();

            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] is null)
                    {
                        continue;
                    }

                    string cellStr = arr[i, j].ToString().Trim();
                    if (cellStr.Length < 3) //配置至少有4个字符.所以,4个以下的都可以跳过
                    {
                        continue; //不用""比较, .Length速度比较快
                    }

                    //var cell = sheet.Cells[i + 1, j + 1];//当单元格值是公式时,没法在configLine里进行add, 因为下面的 nthStr 是 ""

                    if (!cellStr.StartsWith("$tb"))
                    {
                        continue;
                    }

                    string cellPosition = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1); //  {"L15", "付款对象"}, $tb1

                    string nthStr = RegexHelper.GetFirstNumber(cellStr);
                    int nth = Convert.ToInt32(nthStr);
                    if (!bodyConfigCache.ContainsKey(nth))
                    {
                        bodyConfigCache.Add(nth, new EPPlusConfigBodyConfig
                        {
                            Nth = nth,
                            Option = new EPPlusConfigBodyOption()
                            {
                                ConfigLine = new List<EPPlusConfigFixedCell>(),
                                ConfigExtra = new List<EPPlusConfigFixedCell>()
                            }
                        });
                    }

                    var bodyConfig = bodyConfigCache[nth];

                    string cellConfigValue;

                    if (cellStr.StartsWith("$tbs")) //模版摘要/汇总等信息单元格
                    {
                        //string cellConfigValue = Regex.Replace(cellStr, "^[$]tbs" + nthStr, ""); //$需要转义
                        cellConfigValue = cellStr.RemovePrefix($"$tbs{nthStr}").Trim();
                        bodyConfig.Option.ConfigExtra.Add(new EPPlusConfigFixedCell { Address = cellPosition, ConfigValue = cellConfigValue });
                    }
                    else if (cellStr.StartsWith($"$tb{nthStr}$")) //模版提供了多少行,若没有配置,在调用FillData()时默认提供1行  $tb1$1
                    {
                        // string cellConfigValue = Regex.Replace(cellStr, $@"^[$]tb{nth}[$]", ""); //$需要转义, 这个值一般都是数字
                        cellConfigValue = cellStr.RemovePrefix($"$tb{nthStr}$").Trim();
                        if (!int.TryParse(cellConfigValue, out var cellConfigValueInt))
                        {
                            if (string.Compare(cellConfigValue, "max", StringComparison.OrdinalIgnoreCase) == 0) //$tb1$max这种配置的
                            {
                                cellConfigValueInt = EPPlusConfig.MaxRow07 - i;
                            }
                            else
                            {
                                throw new Exception("指定提供了多少行的配置项的值无效");
                            }
                        }

                        if (config.Body[nth].Option.MapperExcelTemplateLine != null)
                        {
                            throw new ArgumentException($"Excel文件中重复配置了项$tb{nthStr}${cellConfigValue}");
                        }

                        config.Body[nth].Option.MapperExcelTemplateLine = cellConfigValueInt;
                    }
                    else if (cellStr.StartsWith($"$tb{nthStr}"))
                    {
                        //string cellConfigValue = Regex.Replace(cellStr, $"^[$]tb{nthStr}", ""); //$需要转义

                        cellConfigValue = cellStr.RemovePrefix($"$tb{nthStr}").Trim();

                        if (sheet.Cells[i + 1, j + 1].Merge)
                        {
                            //只有多行合并时才会影响填充数据时每次新增的行数(多列合并不影响数据的填充)
                            //所以,{"A15:A17", "发生日期"}, 这种情况的key可以写成A15,也可以是A15:K17
                            //{"A15:K17", "发生日期"},只有这种情况的key才要写成A15:K17
                            //后续补充:如果单元格类型是{"A15:A17", "发生日期"},然后key是A15:K17.
                            //在生成excel后打开时会提示***.xlsx中发现不可读取的内容。是否恢复此工作簿的内容.
                            //选择修复文档内容后,里面的内容是正确的(至少我测试的几个是这样的)
                            //所以,同行多列合并的单元格的key 必须是 A15 这种格式的
                            var newKey = sheetMergedCellsList.Find(a => a.Contains(cellPosition));
                            if (newKey is null)
                            {
                                /*描述出现null的情况(经过验证,在EPPlus的4.5.3.2中这个BUG没有了,其他版本不知道)
                                 * 有如下单元格 A2, B2, A3, B3, A4, B4 这6个单元格都有值,
                                 * 然后把这6个单元格给合并起来成一个单元格 (可以用A2:B4来描述)
                                 * 当读取 A2, B2, A3, B3, A3, B4, A2:B4 这7个单元格时,所有的值都为A1
                                 * 因为在合并时,excel会提示 合并单元格时，仅保留左上角的值，而放弃其他值.
                                 * 但是,偏偏有一个神操作会造成 newKey为null. 即:在sheetMergedCellsList中肯定找不到
                                 * 且该操作在excel看上去没问题,但是当程序运行时会让我的下一行代码异常.
                                 * 该操作是用格式刷:把A2:B4个合并后,用格式刷选中D2,让D2:E4合并为一个单元格.
                                 * 此时A2:B4只有A2单元格有值,其他任意单元格都是A2值(取消合并单元格可以验证)
                                 * 但是D2:E4的每个单元格都有值,仅仅是显示为一个单元格.
                                */
                                throw new Exception($"工作簿{sheet.Name}的单元格{cellPosition}存在配置问题,请检查({cellPosition}是合并单元格,请取消合并,并把单元格的值给清空,然后重新合并)");
                            }

                            var cells = newKey.Split(':');

                            if (RegexHelper.GetFirstNumber(cells[0]) == RegexHelper.GetFirstNumber(cells[1])) //是同一行的
                            {
                                bodyConfig.Option.ConfigLine.Add(new EPPlusConfigFixedCell { Address = cellPosition, ConfigValue = cellConfigValue });
                            }
                            else
                            {
                                bodyConfig.Option.ConfigLine.Add(new EPPlusConfigFixedCell { Address = newKey, ConfigValue = cellConfigValue });
                            }
                        }
                        else
                        {
                            bodyConfig.Option.ConfigLine.Add(new EPPlusConfigFixedCell { Address = cellPosition, ConfigValue = cellConfigValue });
                        }
                    }

                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉. 注用这种方式清空值,,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                    sheet.Cells[i + 1, j + 1].Value = null; //修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白(后期已经修复)
                }
            }

            StringBuilder sb = new StringBuilder();
            foreach (var bodyConfig in bodyConfigCache)
            {
                #region 验证

                sb.Clear();
                foreach (var item in bodyConfig.Value.Option.ConfigExtra.GetRepeat(a => new { a.ConfigValue }))
                {
                    sb.Append($@"{item.Address}-{item.ConfigValue},");
                }
                if (sb.RemoveLastChar(',').Length > 0)
                {
                    throw new ArgumentException($"Excel文件中的$tbs{bodyConfig.Key}部分配置了相同的项:{sb}");
                }

                sb.Clear();
                foreach (var item in bodyConfig.Value.Option.ConfigLine.GetRepeat(a => new { a.ConfigValue }))
                {
                    sb.Append($@"{item.Address}-{item.ConfigValue},");
                }
                if (sb.RemoveLastChar(',').Length > 0)
                {
                    throw new ArgumentException($"Excel文件中的$tb{bodyConfig.Key}部分配置了相同的项:{sb}");
                }
                #endregion

                config.Body[bodyConfig.Key].Option.ConfigLine = bodyConfig.Value.Option.ConfigLine;
                config.Body[bodyConfig.Key].Option.ConfigExtra = bodyConfig.Value.Option.ConfigExtra;
            }
        }

        /// <summary>
        /// 设置sheetFoot配置
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startWith"></param>
        /// <returns></returns>
        private static List<EPPlusConfigFixedCell> GetConfigFromExcel(ExcelWorksheet sheet, string startWith)
        {
            if (!startWith.StartsWith("$"))
            {
                throw new ArgumentException("配置项必须是$开头");
            }

            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");

            var fixedCellsInfoList = new List<EPPlusConfigFixedCell>();
            var replaceStr = startWith.RemovePrefix("$");
            for (var i = 0; i < arr.GetLength(0); i++)
            {
                for (var j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] is null)
                    {
                        continue;
                    }

                    string cellStr = arr[i, j].ToString().Trim();
                    if (!cellStr.StartsWith(startWith))
                    {
                        continue;
                    }

                    // {"G6", "公司名称"},
                    string key = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);
                    string val = Regex.Replace(cellStr, $"^[$]{replaceStr}", "").Trim(); //$需要转义
                    if (fixedCellsInfoList.Find(a => a.ConfigValue == val) != null)
                    {
                        throw new ArgumentException($"Excel文件中的{startWith}部分配置了相同的项:{val}");
                    }

                    fixedCellsInfoList.Add(new EPPlusConfigFixedCell() { Address = key, ConfigValue = val.Trim() });
                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉 注用这种方式清空值,,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                    sheet.Cells[i + 1, j + 1].Value = null; //统一用 null 来清空单元格
                }
            }
            return fixedCellsInfoList;
        }

        #endregion

        #region 获得空配置

        public static EPPlusConfig GetEmptyConfig() => new EPPlusConfig()
        {
            Head = new EPPlusConfigFixedCells(),
            Body = new EPPlusConfigBody()
            {
                ConfigList = new List<EPPlusConfigBodyConfig>()
            },
            Foot = new EPPlusConfigFixedCells(),
            Report = new EPPlusReport(),
            IsReport = false,
            DeleteFillDateStartLineWhenDataSourceEmpty = false,
        };

        public static EPPlusConfigSource GetEmptyConfigSource() => new()
        {
            Head = new EPPlusConfigSourceHead(),
            Body = new EPPlusConfigSourceBody(),
            Foot = new EPPlusConfigSourceFoot(),
        };

        #endregion

        #region 模板配置相关

        /// <summary>
        /// 返回模版的 titleLine 和 titleColumn
        /// </summary>
        /// <param name="dataConfigInfo"></param>
        /// <param name="wsIndex"></param>
        /// <param name="wsName"></param>
        /// <param name="titleLine"></param>
        /// <param name="titleColumn"></param>
        public static void GetExcelDataConfigInfo(List<ExcelDataConfigInfo> dataConfigInfo, int wsIndex, string wsName, out int titleLine, out int titleColumn)
        {
            titleLine = 2;
            titleColumn = 1;
            if (dataConfigInfo is null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(wsName))
            {
                var result = dataConfigInfo.Find(info => info.WorkSheetName == wsName);
                if (result != null)
                {
                    titleLine = result.TitleLine;
                    titleColumn = result.TitleColumn;
                    return;
                }
            }

            if (wsIndex > 0)
            {
                var result = dataConfigInfo.Find(info => info.WorkSheetIndex == wsIndex);
                if (result != null)
                {
                    titleLine = result.TitleLine;
                    titleColumn = result.TitleColumn;
                    return;
                }
            }
        }


        /// <summary>
        /// 获得excel填写的配置内容
        /// </summary>
        /// <param name="content"></param>
        /// <param name="outResultPrefix"></param>
        /// <param name="alias"></param>
        /// <returns></returns>
        public static string GetFillDefaultConfig(
            string content,
            string outResultPrefix = "$tb1",
            Dictionary<string, string> alias = null)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }
            alias ??= new Dictionary<string, string>();
            content = content.TrimEnd();
            content.RemoveLastChar('\n');//excel选择列复制出来到文本上有换行,最后一个字符的ascii 是10 \n
            content.RemoveLastChar('\r');//如果是自己敲入的回车,那么也去掉
            var excel_cell_split = new char[] { '	', ' ', };// 两个单元格之间间隔的符号(\t),空格
            string[] splits = content.Split(excel_cell_split, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();
            StringBuilder sbColumn = new StringBuilder();
            foreach (var item in splits)
            {
                var newName = alias.ContainsKey(item)
                    ? NamingHelper.ExtractName(alias[item])
                    : NamingHelper.ExtractName(item);

                sb.Append($@"{outResultPrefix}{newName}{excel_cell_split[0]}");
                sbColumn.AppendLine($"dt.Columns.Add(\"{newName}\");");
            }

            sb.RemoveLastChar(excel_cell_split[0]);

            //sb.AppendLine().AppendLine().AppendLine();
            //sb.AppendLine($@"DataTable dt = new DataTable();");
            //sb.Append(sbColumn.ToString());

            return sb.ToString();
        }

        #endregion

        #region 科学计数法的cell转成文本格式的cell

        /// <summary>
        /// 科学计数法的cell转成文本格式的cell
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="savePath"></param>
        /// <param name="containNoMatchCell">包含不匹配的单元格(即把所有的单元格变成文本格式),true:是.false:仅针对显示成科学计数法的cell变成文本</param>
        /// <returns>是否有进行科学技术法的cell转换.true:是,false:否</returns>
        public static bool ScientificNotationFormatToString(ExcelPackage excelPackage, string savePath, bool containNoMatchCell = false)
        {
            long modifyCellCount = 0;//统计修改的次数

            //下面代码之所以用if-else,是因为这样可以减少很多判断次数.
            if (containNoMatchCell)
            {
                foreach (var sheet in EPPlusHelper.GetExcelWorksheets(excelPackage))
                {
                    object[,] arr = sheet.Cells.Value as object[,];
                    Debug.Assert(arr != null, nameof(arr) + " != null");
                    for (int i = 0; i < arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < arr.GetLength(1); j++)
                        {
                            if (arr[i, j] != null)
                            {
                                modifyCellCount++;
                                sheet.Cells[i + 1, j + 1].Value = sheet.Cells[i + 1, j + 1].Text;
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (var sheet in EPPlusHelper.GetExcelWorksheets(excelPackage))
                {
                    object[,] arr = sheet.Cells.Value as object[,];
                    Debug.Assert(arr != null, nameof(arr) + " != null");
                    for (int i = 0; i < arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < arr.GetLength(1); j++)
                        {
                            if (arr[i, j] != null)
                            {
                                //两段代码唯一的区别
                                var cell = sheet.Cells[i + 1, j + 1];
                                if (cell.Value is double && cell.Text.Length > 11)
                                {
                                    modifyCellCount++;
                                    cell.Value = cell.Text;
                                }
                            }
                        }
                    }
                }
            }

            EPPlusHelper.Save(excelPackage, savePath);

            return modifyCellCount > 0;
        }

        /// <summary>
        /// 处理Excel,将包含科学计数法的cell转成文本格式的cell
        /// </summary>
        /// <param name="fileFullPath">文件路径</param>
        /// <param name="fileSaveAsPath">文件另存为路径</param>
        /// <param name="containNoMatchCell">包含不匹配的单元格(即把所有的单元格变成文本格式),true:是.false:仅针对显示成科学计数法的cell变成文本</param>
        /// <returns>是否有进行科学技术法的cell转换.true:是,false:否</returns>
        public static bool ScientificNotationFormatToString(string fileFullPath, string fileSaveAsPath, bool containNoMatchCell = false)
        {
            using (var fs = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                return ScientificNotationFormatToString(excelPackage, fileSaveAsPath, containNoMatchCell);
            }
        }
        #endregion

        #region GetListErrorMsg

        /// <summary>
        /// 获得错误消息
        /// </summary>
        /// <param name="action"></param>
        /// <returns></returns>
        public static string GetListErrorMsg(Action action)
        {
            try
            {
                action.Invoke();
                return "";
            }
            catch (Exception e)
            {
                return GetListErrorMsg(e);
            }
        }

        /// <summary>
        /// 获得错误消息
        /// </summary>
        /// <param name="action"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static string GetListErrorMsg<T>(Func<List<T>> action, out List<T> result) where T : class, new()
        {
            try
            {
                result = action.Invoke();
                return "";
            }
            catch (Exception e)
            {
                result = new List<T>();
                return GetListErrorMsg(e);
            }
        }

        private static string GetListErrorMsg(Exception e)
        {
            StringBuilder sb = new StringBuilder();
            if (e?.Message?.Length > 0)
            {
                sb.AppendLine("程序报错:Message:");
                sb.Append(e.Message);
            }
            if (e?.InnerException?.Message?.Length > 0)
            {
                sb.AppendLine("程序报错:InnerExceptionMessage:");
                sb.Append(e.InnerException.Message);
            }
            var txt = sb.ToString();
            return txt;
        }
        #endregion

        #region 文件相关的帮助方法

        /// <summary>
        /// 读取一个文件,获得一个文件流
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static FileStream GetFileStream(string filePath)
        {
            FileMode mode = FileMode.Open;
            FileAccess access = FileAccess.Read;
            FileShare share = FileShare.ReadWrite;
            return FileHelper.GetFileStream(filePath, mode, access, share);
        }

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="filePath">文件路径</param>
        public static void Save(ExcelPackage excelPackage, string filePath)
        {
            ExcelPackageHelper.Save(excelPackage, filePath);
        }

        #endregion

    }
}