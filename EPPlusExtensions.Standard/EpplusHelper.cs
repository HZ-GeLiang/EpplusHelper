using EPPlusExtensions.Attributes;
using EPPlusExtensions.Helper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using EPPlusExtensions.Exceptions;

namespace EPPlusExtensions
{
    public class EPPlusHelper
    {
        /// <summary>
        /// 填充Excel时创建的工作簿名字
        /// </summary>
        public static List<string> FillDataWorkSheetNameList = new List<string>();

        public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; //类型参考网址: http://filext.com/faq/office_mime_types.php

        #region GetExcelWorksheet

        /// <summary>
        /// 获得当前Excel的所有工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
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
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            workSheetIndex = ConvertWorkSheetIndex(excelPackage, workSheetIndex);
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
            if (copyWorkSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(copyWorkSheetIndex));
            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
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
            ws.Name = workSheetNewName;
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

        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string workName, bool onlyOneWorkSheetReturnThis)
        {
            ExcelWorksheet ws = null;
            if (onlyOneWorkSheetReturnThis && excelPackage.Workbook.Worksheets.Count == 1)
            {
                var workSheetIndex = ConvertWorkSheetIndex(excelPackage, 1);
                ws = excelPackage.Workbook.Worksheets[workSheetIndex];
            }
            if (ws != null)
            {
                return ws;
            }
            if (workName == null) throw new ArgumentNullException(nameof(workName));
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
        /// <param name="destWorkSheetName"></param>
        /// <param name="workSheetNewName"></param>
        /// <returns></returns>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string destWorkSheetName, string workSheetNewName)
        {
            if (destWorkSheetName == null) throw new ArgumentNullException(nameof(destWorkSheetName));
            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
            var wsMom = GetExcelWorksheet(excelPackage, destWorkSheetName);
            var ws = excelPackage.Workbook.Worksheets.Add(workSheetNewName, wsMom);
            ws.Name = workSheetNewName;
            return ws;
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
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));

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

            workSheetIndex = ConvertWorkSheetIndex(excelPackage, workSheetIndex);
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
            if (eWorkSheetHiddens == null) return;
            if (workSheetNameExcludeList == null) workSheetNameExcludeList = new List<string>();
            var delWsNames = GetWorkSheetNames(excelPackage, eWorkSheetHiddens);
            foreach (var wsName in delWsNames)
            {
                if (workSheetNameExcludeList.Contains(wsName)) continue;
                EPPlusHelper.DeleteWorksheet(excelPackage, wsName);
            }
        }

        /// <summary>
        /// 获得工作簿,根据第二个参数,可以用来 获得隐藏的工作簿
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="eWorkSheetHiddens"></param>
        /// <returns></returns>
        public static List<string> GetWorkSheetNames(ExcelPackage excelPackage, params eWorkSheetHidden[] eWorkSheetHiddens)
        {
            var wsNames = new List<string>();
            if (eWorkSheetHiddens == null || eWorkSheetHiddens.Length == 0) return wsNames;

            for (int i = 1; i <= excelPackage.Workbook.Worksheets.Count; i++)
            {
                var index = ConvertWorkSheetIndex(excelPackage, i);
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
        public static void DeleteWorksheetAll(ExcelPackage excelPackage, List<string> workSheetNameExcludeList)
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
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么.  </param>
        /// <param name="destWorkSheetName">填充数据的workSheet叫什么</param> 
        public static void FillData(ExcelPackage excelPackage, EPPlusConfig config, EPPlusConfigSource configSource, string workSheetNewName, string destWorkSheetName)
        {
            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
            if (destWorkSheetName == null) throw new ArgumentNullException(nameof(destWorkSheetName));
            ExcelWorksheet worksheet = GetExcelWorksheet(excelPackage, destWorkSheetName, workSheetNewName);
            EPPlusHelper.FillDataWorkSheetNameList.Add(workSheetNewName);
            config.WorkSheetDefault?.Invoke(worksheet);
            EPPlusHelper.FillData(config, configSource, worksheet);
        }

        /// <summary>
        /// 往目标sheet中填充数据.注:填充的数据的类型会影响你设置的excel单元格的格式是否起作用
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么. 若为""/null,则默认是"Sheet" +workSheetNewName </param>
        /// <param name="destWorkSheetIndex">从1开始</param>
        public static void FillData(ExcelPackage excelPackage, EPPlusConfig config, EPPlusConfigSource configSource, string workSheetNewName, int destWorkSheetIndex)
        {
            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
            if (destWorkSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(destWorkSheetIndex));

            ExcelWorksheet worksheet = EPPlusHelper.DuplicateWorkSheetAndRename(excelPackage, destWorkSheetIndex, workSheetNewName);
            EPPlusHelper.FillDataWorkSheetNameList.Add(workSheetNewName);//往list里添加数据
            config.WorkSheetDefault?.Invoke(worksheet);
            EPPlusHelper.FillData(config, configSource, worksheet);
        }

        /// <summary>
        /// 往目标sheet中填充数据
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        private static void FillData(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            EPPlusHelper.FillData_Head(config, configSource, worksheet);
            int sheetBodyAddRowCount = 0;
            if (configSource?.Body?.ConfigList.Count > 0)
            {
                long allDataTableRows = 0;
                foreach (var bodyInfo in configSource.Body.ConfigList)
                {
                    allDataTableRows += bodyInfo?.Option?.DataSource?.Rows.Count ?? 0;
                }
                //这个限制仅仅针对,标题行是单行, 且填充数据是单行,没有FillData_Head 和 FillData_Foot 才有效
                if (allDataTableRows > EPPlusConfig.MaxRow07 - 1)//-1是去掉第一行的标题行
                {
                    throw new IndexOutOfRangeException("要导出的数据行数超过excel最大行限制");
                }

                sheetBodyAddRowCount = EPPlusHelper.FillData_Body(config, configSource, worksheet);
            }

            EPPlusHelper.FillData_Foot(config, configSource, worksheet, sheetBodyAddRowCount);
        }


        #endregion

        #region 私有方法

        private static void FillData_Head(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            //填充head
            if (config.Head.ConfigCellList == null || config.Head.ConfigCellList.Count <= 0)
            {
                return;
            }

            var dictConfigSourceHead = configSource.Head.CellsInfoList.ToDictionary(a => a.ConfigValue);

            foreach (var item in config.Head.ConfigCellList)
            {
                if (configSource.Head == null || configSource.Head.CellsInfoList == null ||
                    configSource.Head.CellsInfoList.Count <= 0) //excel中有配置head,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
                {
                    break;
                }

                string colMapperName = item.ConfigValue;
                object val = config.Head.ConfigItemMustExistInDataColumn
                    ? dictConfigSourceHead[item.ConfigValue].FillValue
                    : dictConfigSourceHead.ContainsKey(item.ConfigValue) ? dictConfigSourceHead[item.ConfigValue].FillValue : null;

                ExcelRange cells = worksheet.Cells[item.Address];

                if (config.Head.CellCustomSetValue != null)
                {
                    config.Head.CellCustomSetValue.Invoke(colMapperName, val, cells);
                }
                else
                {
                    SetWorksheetCellsValue(config, cells, val, colMapperName);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        /// <returns>新增了多少行</returns>
        private static int FillData_Body(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            //填充body
            var sheetBodyAddRowCount = 0; //新增了几行 (统计sheet body 在原有的模版上新增了多少行), 需要返回的

            if (config == null || configSource == null ||
                config.Body == null || configSource.Body == null ||
                config.Body.ConfigList == null || configSource.Body.ConfigList == null ||
                config.Body.ConfigList.Count <= 0 || configSource.Body.ConfigList.Count <= 0)
            {
                return sheetBodyAddRowCount;
            }

            int sheetBodyDeleteRowCount = 0; //sheet body 中删除了多少行(只含配置的行,对于FillData()内的删除行则不包括在内).  
            var dictConfig = config.Body.ConfigList.ToDictionary(a => a.Nth, a => a.Option);
            var dictConfigSource = configSource.Body.ConfigList.ToDictionary(a => a.Nth, a => a.Option);
            foreach (var itemInfo in config.Body.ConfigList)
            {
                var nth = itemInfo.Nth;//body的第N个配置

                #region get dataTable
                DataTable datatable;
                if (!dictConfigSource.ContainsKey(nth)) //如果没有数据源中没有excle中配置
                {
                    //需要删除配置行(当数据源为空[无,null.rows.count=0])
                    if (!config.DeleteFillDateStartLineWhenDataSourceEmpty)
                    {
                        continue; //不需要删除删除配置空行,那么直接跳过
                    }
                    datatable = null;
                }
                else
                {
                    datatable = dictConfigSource[nth].DataSource; //body的第N个配置的数据源
                }

                #endregion

                #region When dataTable is empty
                if (datatable == null || datatable.Rows.Count <= 0) //数据源为null或为空
                {
                    //throw new ArgumentNullException($"configSource.SheetBody[{nth.Key}]没有可读取的数据");

                    if (!config.DeleteFillDateStartLineWhenDataSourceEmpty || dictConfig[nth].ConfigLine.Count <= 0)
                    {
                        continue; //跳过本次fillDate的循环
                    }

                    #region DeleteFillDateStartLine

                    foreach (var cellConfigInfo in dictConfig[nth].ConfigLine) //只遍历一次
                    {
                        int driftVale = 1; //浮动值,如果是合并单元格,则取合并单元格的行数
                        int delRow; //要删除的行号
                        if (cellConfigInfo.Address.Contains(":")) //如果是合并单元格,修改浮动的行数
                        {
                            var cells = cellConfigInfo.Address.Split(new[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                            if (cells.Length != 2) throw new Exception("该合并单元格的标识有问题,不是类似于A1:A2这个格式的");
                            int mergeCellStartRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[0]));
                            int mergeCellEndRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[1]));

                            driftVale = mergeCellEndRow - mergeCellStartRow + 1;
                            if (driftVale <= 0) throw new Exception("合并单元格的合并行数小于1");

                            delRow = mergeCellStartRow + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }
                        else //不是合并单元格
                        {
                            delRow = Convert.ToInt32(RegexHelper.GetLastNumber(cellConfigInfo.Address)) + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }

                        if (delRow <= 0) throw new Exception("要删除的行号不能小于0");

                        worksheet.DeleteRow(delRow, driftVale, true);
                        sheetBodyDeleteRowCount += driftVale;
                        break; //只要删一次即可
                    }

                    #endregion

                    continue; //强制跳过本次fillDate的循环
                }
                #endregion

                int currentLoopAddLines = 0;
                var deleteLastSpaceLine = false; //是否删除最后一空白行(可能有多行组成的)
                int lastSpaceLineInterval = 0; //表示最后一空白行由多少行组成,默认为0
                int lastSpaceLineRowNumber = 0; //表示最后一行的行号是多少
                int tempLine = dictConfig[nth].MapperExcelTemplateLine ?? 1; //获得第N个配置中excel模版提供了多少行,默认1行
                var hasMergeCell = dictConfig[nth].ConfigLine.Find(a => a.Address.Contains(":")) != null;
                Dictionary<string, FillDataColumns> fillDataColumnsStat = null;//Datatable 的列的使用情况  

                //3.赋值
                var customValue = new CustomValue
                {
                    ConfigLine = config.Body[nth].Option.ConfigLine,
                    ConfigExtra = config.Body[nth].Option.ConfigExtra,
                    Worksheet = worksheet,
                };  //注: 这里没有用深拷贝,所以,在使用的时候,不要修改内部的值, 否则后果自负.

                if (hasMergeCell)
                {
                    //注:进入这里的条件是单元格必须是多行合并的,如果是同行多列合并的单元格,最后生成的excel会有问题,打开时会提示修复(修复完成后内容是正确的(不保证,因为我测试的几个内容是正确的))
                    List<ExcelCellRange> cellRange = dictConfig[nth].ConfigLine.Select(cellConfigInfo => new ExcelCellRange(cellConfigInfo.Address)).ToList();
                    int maxIntervalRow = (from c in cellRange select c.IntervalRow).Max();
                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        DataRow row = datatable.Rows[i];
                        int destRow;

                        if (nth == 1)
                        {
                            destRow = cellRange[0].Start.Row + i * (maxIntervalRow + 1) - sheetBodyDeleteRowCount;
                        }
                        else
                        {
                            destRow = currentLoopAddLines > 0
                                ? cellRange[0].Start.Row + (tempLine - 1) * (maxIntervalRow + 1) + sheetBodyAddRowCount -
                                  sheetBodyDeleteRowCount
                                : cellRange[0].Start.Row + i * (maxIntervalRow + 1) + sheetBodyAddRowCount -
                                  sheetBodyDeleteRowCount;
                        }

                        if (datatable.Rows.Count > 1) //1.数据源中的数据行数大于1才增行
                        {
                            //1.新增一行
                            if (i > tempLine - 2) //i从0开始,这边要-1,然后又要留一行模版,做为复制源,所以这里-2
                            {
                                if (i == tempLine - 1) //仅剩余最后一行是空白的
                                {
                                    deleteLastSpaceLine = true;
                                    lastSpaceLineInterval = maxIntervalRow + 1;
                                }

                                lastSpaceLineRowNumber = destRow + maxIntervalRow + 1; //最后一行空行的位置
                                worksheet.InsertRow(destRow, maxIntervalRow + 1, destRow + maxIntervalRow + 1); //新增N行,注意,此时新增行的高度是有问题的
                                                                                                                //2.复制样式(含修正)
                                for (int j = 0; j <= maxIntervalRow; j++) //修正height
                                {
                                    worksheet.Row(destRow + j).Height = worksheet.Row(destRow + j + maxIntervalRow + 1).Height;
                                }

                                sheetBodyAddRowCount += maxIntervalRow + 1;
                                currentLoopAddLines++;
                            }
                        }

                        //3.赋值
                        for (int j = 0; j < dictConfig[nth].ConfigLine.Count; j++)//遍历列
                        {
                            #region 赋值
                            string colMapperName = dictConfig[nth].ConfigLine[j].ConfigValue;
                            object val = dictConfig[nth].ConfigItemMustExistInDataColumn
                                ? row[colMapperName]
                                : row.Table.Columns.Contains(colMapperName) ? row[colMapperName] : null;
#if DEBUG
                            if (!cellRange[j].IsMerge)
                            {
                                throw new Exception("填充数据时,合并单元格填充处不是合并单元格,请修改组件代码");
                            }
#endif

                            //int destRowStart = cellRange[j].Start.Row;
                            int destStartCol = cellRange[j].Start.Col;
                            //int destEndRow = cellRange[j].End.Row;
                            int destEndCol = cellRange[j].End.Col;
                            if (!worksheet.Cells[destRow, destStartCol, destRow + maxIntervalRow, destEndCol].Merge)
                            {
                                /*
                                 注:假设原有的cell单元格是同行多列合并,那么上面的if判断还是会返回false.(但在worksheet.MergedCells中能发现单元格是合并的).多列合并的没试过.
                                 后来做了限制,进入这个if语句内的单元格必须是多行合并的,对于单行多列合并的cell,后果自负
                                 */
                                worksheet.Cells[destRow, destStartCol, destRow + maxIntervalRow, destEndCol].Merge = true;
                            }

                            //worksheet.Cells[destRow, destStartCol, destRow + maxIntervalRow, destEndCol].Value = val;

                            var cells = worksheet.Cells[destRow, destStartCol, destRow + maxIntervalRow, destEndCol];

                            if (dictConfig[nth].CustomSetValue != null)
                            {
                                customValue.Area = FillArea.Content;
                                customValue.ColName = colMapperName;
                                customValue.Value = val;
                                customValue.Cell = cells;
                                dictConfig[nth].CustomSetValue.Invoke(customValue);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }
                            #endregion
                        }

                        if (config.IsReport)
                        {
                            SetReport(worksheet, row, config, destRow, maxIntervalRow);
                        }
                    }
                }
                else //sheet body是常规类型的,即,没有合并单元格的(或者是同行多列的单元格)
                {
                    var configLineCellPoint = dictConfig[nth].ConfigLine.Select(configCell => new ExcelCellPoint(configCell.Address)).ToList(); // 将配置的值 转换成 ExcelCellPoint
                    var leftCellInfo = configLineCellPoint.First();

                    #region 必须在 insertRow 之前计算,否则,当前变量就是插入行后的单元格信息
                    var rightCellInfo = configLineCellPoint.Last();
                    var leftColStr = leftCellInfo.ColStr;
                    var rightColStr = worksheet.Cells[rightCellInfo.R1C1].Merge
                        ? new ExcelCellRange(rightCellInfo.R1C1, worksheet).End.ColStr
                        : rightCellInfo.ColStr;
                    #endregion

                    #region 第一遍循环:计算要插入多少行

                    var insertRows = 0;//要新增多少行
                    var insertRowFrom = 0;//从哪行开始
                    var dictDestRow = new Dictionary<int, int>();//数据源的第N行,对应excel填充的第N行
                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        int destRow = nth == 1
                            ? sheetBodyAddRowCount > 0
                                ? leftCellInfo.Row + i - sheetBodyDeleteRowCount
                                : leftCellInfo.Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                            : currentLoopAddLines > 0
                                ? leftCellInfo.Row + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                                : leftCellInfo.Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;

                        dictDestRow.Add(i, destRow);

                        if (datatable.Rows.Count <= 1) continue; //1.数据源中的数据行数大于1才增行
                        if (i <= tempLine - 2) continue; //i从0开始,这边要-1,然后又要留一行模版,做为复制源,所以这里要-2  
                        if (i == tempLine - 1) //仅剩余最后一行是空白的
                        {
                            deleteLastSpaceLine = true;
                            lastSpaceLineInterval = 1;
                        }
                        lastSpaceLineRowNumber = destRow + 1; //最后一行空行的位置
                        if (insertRowFrom == 0)
                        {
                            insertRowFrom = destRow;
                        }
                        insertRows++;
                        sheetBodyAddRowCount++;
                        currentLoopAddLines++;
                    }
                    #endregion

                    var needInsert = insertRows > 0 && insertRowFrom > 0;
                    if (needInsert)
                    {
                        //在  InsertRowFrom 行前面插入 InsertRowCount 行.
                        //注:
                        //1. 新增的行的Height的默认值,需要自己修改
                        //2. copyStylesFromRow 的行计算是在 InsertRowFrom+ InsertRowCount 后开始的那行.
                        //3. copyStylesFromRow 不会把合并的单元格也弄过来(即,这个参数的功能不是格式刷)
                        if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyAll)
                        {
                            worksheet.InsertRow(insertRowFrom, insertRows); //用这个参数创建的excel,文件体积要小,插入速度没测试
                        }
                        else if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyStyleAndMergeCell)
                        {
                            if (dictConfig[nth].InsertRowStyle.NeedCopyStyles)
                            {
                                //在测试中,数据量 >= EPPlusConfig.MaxRow07/2-1  时,程序会抛异常, 这个数据量值仅做参考
                                //解决方案,分批插入, 且分批插入的 RowFrom 必须是第一次 InsertRow 的结尾行, 不然 '第三遍循环:填充数据' 会异常
                                //同时又发现了一个bug: worksheet.InsertRow() 的第三个参数要满足 _rows + copyStylesFromRow < EPPlusConfig.MaxRow07 , 但是_copyStylesFromRow 又是 rowFrom + rows 后开始数的行数
                                //nnd. 为了 屏蔽这个bug报错, 我后面写了if-else.  这样写的 结果就是 后面新增的行没有样式

                                var insertRowsMax = (EPPlusConfig.MaxRow07 / 2 - 1) - 1;
                                if (insertRows >= insertRowsMax)
                                {
                                    var insertCount = insertRows / insertRowsMax;
                                    var mod = insertRows % insertRowsMax;
                                    int rowFrom; int rows; int copyStylesFromRow;
                                    for (int i = 0; i < insertCount; i++)
                                    {
                                        rowFrom = insertRowFrom + i * insertRowsMax;
                                        rows = insertRowsMax;
                                        copyStylesFromRow = rowFrom + rows;
                                        //防止报错, 结果就是 后面新增的行没有样式
                                        if (rows + copyStylesFromRow > EPPlusConfig.MaxRow07)
                                        {
                                            worksheet.InsertRow(rowFrom, rows);
                                        }
                                        else
                                        {
                                            worksheet.InsertRow(rowFrom, rows, copyStylesFromRow);
                                        }
                                    }
                                    if (mod > 0)
                                    {
                                        rowFrom = insertRowFrom + insertCount * insertRowsMax;
                                        rows = mod;
                                        copyStylesFromRow = lastSpaceLineRowNumber;
                                        //防止报错, 结果就是 后面新增的行没有样式
                                        if (rows + copyStylesFromRow > EPPlusConfig.MaxRow07)
                                        {
                                            worksheet.InsertRow(rowFrom, rows);
                                        }
                                        else
                                        {
                                            worksheet.InsertRow(rowFrom, rows, copyStylesFromRow);
                                        }

                                    }
                                }
                                else
                                {
                                    worksheet.InsertRow(insertRowFrom, insertRows, lastSpaceLineRowNumber);
                                }
                            }
                            else
                            {
                                worksheet.InsertRow(insertRowFrom, insertRows);
                            }
                        }
                    }

                    #region 第二遍循环:处理样式 (Height要自己单独处理)
                    if (needInsert)
                    {
                        if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyAll)
                        {
                            var configLine = $"{leftColStr}{lastSpaceLineRowNumber}:{rightColStr}{lastSpaceLineRowNumber}";

                            for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                            {
                                int destRow = dictDestRow[i];

                                //copy 好比格式刷, 这里只格式化配置行所在的表格部分.
                                //Copy 效率比 CopyStyleAndMergedCellFromConfigRow 慢差不多一倍(测试数据4w条,要4秒多, 用上面的是2秒多,且文件体积也要小很多 好像有50% ) 
                                worksheet.Cells[configLine].Copy(worksheet.Cells[$"{leftColStr}{destRow}:{rightColStr}{destRow}"]);//注: 如果rightColStr后还有单元格,请参考Sample05

                                //不要用[row,col]索引器,[row,col]表示某单元格.注意:copy会把source行的除了height(觉得是一个bug)以外的全部复制一行出来
                                worksheet.Row(destRow).Height = worksheet.Row(lastSpaceLineRowNumber).Height; //修正height
                            }
                        }
                        else if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyStyleAndMergeCell)
                        {
                            var modifyInsertRowHeight = true;
                            if (dictConfig[nth].InsertRowStyle.NeedMergeCell)
                            {
                                List<ExcelCellRange> rangeCells = GetMergedCellFromRow(worksheet, lastSpaceLineRowNumber, leftColStr, rightColStr);
                                if (rangeCells != null && rangeCells.Count > 0)
                                {
                                    modifyInsertRowHeight = false;
                                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                                    {
                                        int destRow = dictDestRow[i];
                                        foreach (var item in rangeCells)
                                        {
                                            worksheet.Cells[destRow, item.Start.Col, destRow, item.End.Col].Merge = true;
                                        }
                                        worksheet.Row(destRow).Height = worksheet.Row(lastSpaceLineRowNumber).Height; //修正height
                                    }
                                }
                            }
                            if (modifyInsertRowHeight)
                            {
                                for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                                {
                                    int destRow = dictDestRow[i];
                                    worksheet.Row(destRow).Height = worksheet.Row(lastSpaceLineRowNumber).Height; //修正height
                                }
                            }
                        }
                    }
                    #endregion

                    #region 第三遍循环:填充数据

                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        int destRow = dictDestRow[i];
                        DataRow row = datatable.Rows[i];

                        //3.赋值.
                        //注:遍历时变量 j 的终止条件不能是 dataTable.Rows.Count. 因为dataTable可能会包含多余的字段信息,与 配置信息列的个数不一致.
                        for (int j = 0; j < dictConfig[nth].ConfigLine.Count; j++)
                        {
                            #region 赋值

                            //worksheet.Cells[destRow, destCol].Value = row[j];
                            string colMapperName = dictConfig[nth].ConfigLine[j].ConfigValue;//身份证

                            if (string.IsNullOrEmpty(colMapperName))
                            {
                                continue;
                            }

                            //33xxxx19941111xxxx
                            object val = dictConfig[nth].ConfigItemMustExistInDataColumn
                                ? row[colMapperName]
                                : row.Table.Columns.Contains(colMapperName) ? row[colMapperName] : null;
                            int destCol = configLineCellPoint[j].Col;
                            var cells = worksheet.Cells[destRow, destCol];

                            if (dictConfig[nth].CustomSetValue != null)
                            {
                                customValue.Area = FillArea.Content;
                                customValue.ColName = colMapperName;
                                customValue.Value = val;
                                customValue.Cell = cells;
                                dictConfig[nth].CustomSetValue.Invoke(customValue);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }

                            #endregion

                            #region 同步数据源

                            if (j == configLineCellPoint.Count - 1) //如果一行循环到了最后一列
                            {
                                if (dictConfigSource[nth].FillMethod == null)
                                {
                                    continue;
                                }
                                var fillMethod = dictConfigSource[nth].FillMethod;
                                if (fillMethod == null || fillMethod.FillDataMethodOption == SheetBodyFillDataMethodOption.Default)
                                {
                                    continue;
                                }
                                if (fillMethod.FillDataMethodOption == SheetBodyFillDataMethodOption.SynchronizationDataSource)
                                {
                                    var isFillDataTitle = fillMethod.SynchronizationDataSource.NeedTitle && i == 0;
                                    var isFillDataBody = fillMethod.SynchronizationDataSource.NeedBody;

                                    if (!isFillDataTitle && !isFillDataBody) continue;

                                    if (fillDataColumnsStat == null)
                                    {
                                        fillDataColumnsStat = InitFillDataColumnStat(datatable, dictConfig[nth].ConfigLine, fillMethod);
                                    }

                                    if (isFillDataTitle)
                                    {
                                        var eachCount_Col = 0;
                                        var lastConfigCell = dictConfig[nth].ConfigLine.Last();
                                        var eachCount_Col_Step = lastConfigCell.IsMergeCell == true
                                            ? new ExcelCellRange(lastConfigCell.FullAddress).IntervalCol + 1
                                            : 1;

                                        var config_firstCell_Col = new ExcelCellPoint(dictConfig[nth].ConfigLine.First().Address).Col;
                                        var config_lastCell_col = new ExcelCellPoint(lastConfigCell.Address).Col;
                                        var titleLine_LastCell = worksheet.Cells[destRow - 1, config_lastCell_col];//标题行的最后一列的address

                                        foreach (var item in fillDataColumnsStat.Values)
                                        {
                                            if (item.State != FillDataColumnsState.WillUse) continue;
                                            var extensionDestCol_title_Col = config_firstCell_Col + dictConfig[nth].ConfigLineInterval + eachCount_Col;
                                            var extensionCell_title = worksheet.Cells[destRow - 1, extensionDestCol_title_Col];
                                            extensionCell_title.StyleID = titleLine_LastCell.StyleID;

                                            if (dictConfig[nth].CustomSetValue != null)
                                            {
                                                customValue.Area = FillArea.TitleExt;
                                                customValue.ColName = item.ColumnName;
                                                customValue.Value = item.ColumnName;
                                                customValue.Cell = extensionCell_title;
                                                dictConfig[nth].CustomSetValue.Invoke(customValue);
                                            }
                                            else
                                            {
                                                SetWorksheetCellsValue(config, extensionCell_title, item.ColumnName, item.ColumnName);
                                            }
                                            eachCount_Col += eachCount_Col_Step;
                                        }
                                    }

                                    if (isFillDataBody)
                                    {
                                        var eachCount_Col = 0;
                                        var lastConfigCell = dictConfig[nth].ConfigLine.Last();
                                        var eachCount_Col_Step = lastConfigCell.IsMergeCell == true
                                            ? new ExcelCellRange(lastConfigCell.FullAddress).IntervalCol + 1
                                            : 1;

                                        var lastCell_IntervalCol = eachCount_Col_Step - 1;
                                        var config_lastCell_Col = new ExcelCellPoint(lastConfigCell.Address).Col;//配置列的最后一个address
                                        var lastCell = worksheet.Cells[destRow, config_lastCell_Col];

                                        foreach (var item in fillDataColumnsStat.Values)
                                        {
                                            if (item.State != FillDataColumnsState.WillUse) continue;
                                            var extensionDestCol_body_Col = (configLineCellPoint[j].Col + 1) + eachCount_Col + lastCell_IntervalCol;
                                            var extensionCell_body = worksheet.Cells[destRow, extensionDestCol_body_Col];
                                            extensionCell_body.StyleID = lastCell.StyleID;

                                            //还有好多样式没有弄
                                            //8.设置字体
                                            //extensionCell_body.Style.Font.xxx =..
                                            //9.设置边框的属性
                                            //extensionCell_body.Style.Border.Left.Style = ...
                                            //extensionCell_body.Style.Border.Right.Style = ...
                                            //extensionCell_body.Style.Border.Top.Style = ...
                                            //extensionCell_body.Style.Border.Bottom.Style = ...
                                            //10.对齐方式
                                            //extensionCell_body.HorizontalAlignment = ...
                                            //extensionCell_body.VerticalAlignment = ...
                                            //11.设置整个sheet的背景色
                                            //extensionCell_body.Fill.PatternType = ...
                                            //extensionCell_body.Fill.BackgroundColor.SetColor(...);

                                            SetWorksheetCellsValue(config, extensionCell_body, row[item.ColumnName], item.ColumnName);
                                            if (dictConfig[nth].CustomSetValue != null)
                                            {
                                                customValue.Area = FillArea.ContentExt;
                                                customValue.ColName = item.ColumnName;
                                                customValue.Value = row[item.ColumnName];
                                                customValue.Cell = extensionCell_body;
                                                dictConfig[nth].CustomSetValue.Invoke(customValue);
                                            }
                                            else
                                            {
                                                SetWorksheetCellsValue(config, extensionCell_body, row[item.ColumnName], item.ColumnName);
                                            }
                                            eachCount_Col += eachCount_Col_Step;
                                        }
                                    }
                                }
                            }

                            #endregion

                        }

                        if (config.IsReport)
                        {
                            SetReport(worksheet, row, config, destRow);
                        }
                    }
                    #endregion
                }

                //已经修复bug:当只有一个配置时,且 deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白 原因是,配置行没被删除
                if (deleteLastSpaceLine)
                {
                    worksheet.DeleteRow(lastSpaceLineRowNumber, lastSpaceLineInterval, true);
                    sheetBodyAddRowCount -= lastSpaceLineInterval;
                }

                #region FillData_Body_Summary
                //填充第N个配置的一些零散的单元格的值(譬如汇总信息等)

                if (dictConfigSource[nth].ConfigExtra != null)
                {
                    var dictConfigSourceSummary = dictConfigSource[nth].ConfigExtra.Source.ToDictionary(a => a.ConfigValue);
                    foreach (var item in dictConfig[nth].ConfigExtra)
                    {
                        var excelCellPoint = new ExcelCellPoint(item.Address);
                        string colMapperName = item.ConfigValue;
                        object val = dictConfigSourceSummary[colMapperName].FillValue;
                        ExcelRange cells = worksheet.Cells[excelCellPoint.Row + sheetBodyAddRowCount, excelCellPoint.Col];

                        if (dictConfig[nth].SummaryCustomSetValue != null)
                        {
                            customValue.Area = null;
                            customValue.ColName = colMapperName;
                            customValue.Value = val;
                            customValue.Cell = cells;

                            dictConfig[nth].SummaryCustomSetValue.Invoke(customValue);
                        }
                        else
                        {
                            SetWorksheetCellsValue(config, cells, val, colMapperName);
                        }
                    }
                }
                #endregion
            }

            return sheetBodyAddRowCount;
        }

        /// <summary>
        /// 所有的合并单元格
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="lineNo">行号</param>
        /// <param name="leftCol">最左边的</param>
        /// <param name="rightCol">最右边的,如果最右边的合并单元格,取合并单元格的最右边列的地址</param>
        /// <returns></returns>
        private static List<ExcelCellRange> GetMergedCellFromRow(ExcelWorksheet worksheet, int lineNo, string leftCol, string rightCol)
        {
            var allCell = GetCellFromRow(worksheet, lineNo, leftCol, rightCol);

            var rangeCells = new List<ExcelCellRange>();
            foreach (var item in allCell)
            {
                if (item is ExcelCellRange cellRange && cellRange.IsMerge)
                {
                    rangeCells.Add(cellRange);
                }
            }

            return rangeCells;
        }

        /// <summary>
        /// 所有的单元格
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="lineNo">行号</param>
        /// <param name="leftCol">最左边的</param>
        /// <param name="rightCol">最右边的,如果最右边的合并单元格,取合并单元格的最右边列的地址</param>
        /// <returns></returns>
        private static List<object> GetCellFromRow(ExcelWorksheet worksheet, int lineNo, string leftCol, string rightCol)
        {
            var leftAddressCol = ExcelCellPoint.R1C1Formulas(leftCol);
            var rightAddressCol = ExcelCellPoint.R1C1Formulas(rightCol);

            var allCell = new List<object>();
            while (true)
            {
                if (EPPlusHelper.IsMergeCell(worksheet, row: lineNo, col: leftAddressCol, out var mergeCellAddress))
                {
                    var cell = new ExcelCellRange(mergeCellAddress);
                    allCell.Add(cell);
                    leftAddressCol = cell.End.Col + 1;
                }
                else
                {
                    var cellAddress = new ExcelCellPoint(lineNo, leftAddressCol).R1C1;
                    var cell = new ExcelCellPoint(cellAddress);
                    allCell.Add(cell);
                    leftAddressCol++;
                }

                if (leftAddressCol > rightAddressCol)
                {
                    break;
                }
            }
            return allCell;

        }

        /// <summary>
        ///  获得Database数据源的所有的列的使用状态
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="configLine"></param>
        /// <param name="fillModel"></param>
        /// <returns></returns>
        private static Dictionary<string, FillDataColumns> InitFillDataColumnStat(DataTable dataTable, List<EPPlusConfigFixedCell> configLine, SheetBodyFillDataMethod fillModel)
        {
            var fillDataColumnStat = new Dictionary<string, FillDataColumns>();
            foreach (DataColumn column in dataTable.Columns)
            {
                fillDataColumnStat.Add(column.ColumnName, new FillDataColumns()
                {
                    ColumnName = column.ColumnName,
                    State = FillDataColumnsState.Unchanged
                });
            }

            foreach (var item in configLine)
            {
                fillDataColumnStat[item.ConfigValue].State = FillDataColumnsState.Used;
            }

            var isEmptyInclude = string.IsNullOrEmpty(fillModel.SynchronizationDataSource.Include);
            var isEmptyExclude = string.IsNullOrEmpty(fillModel.SynchronizationDataSource.Exclude);
            if (isEmptyInclude != isEmptyExclude) //只能有一个值有效
            {
                if (!isEmptyInclude)
                {
                    Modify_DataColumnsIsUsedStat(fillDataColumnStat, fillModel.SynchronizationDataSource.Include, true);
                }

                if (!isEmptyExclude)
                {
                    Modify_DataColumnsIsUsedStat(fillDataColumnStat, fillModel.SynchronizationDataSource.Exclude, false);
                }
            }

            return fillDataColumnStat;
        }

        private static void Modify_DataColumnsIsUsedStat(Dictionary<string, FillDataColumns> fillDataColumnsStat, string columns, bool selectColumnIsWillUse)
        {
            var splitInclude = columns.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var key in fillDataColumnsStat.Keys)
            {
                if (fillDataColumnsStat[key].State == FillDataColumnsState.Unchanged)
                {
                    if (splitInclude.Contains(key))
                    {
                        fillDataColumnsStat[key].State = selectColumnIsWillUse ? FillDataColumnsState.WillUse : FillDataColumnsState.WillNotUse;
                    }
                    else
                    {
                        fillDataColumnsStat[key].State = selectColumnIsWillUse ? FillDataColumnsState.WillNotUse : FillDataColumnsState.WillUse;
                    }
                }
            }
        }

        /// <summary>
        /// 填充foot
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        /// <param name="sheetBodyAddRowCount"></param>
        private static void FillData_Foot(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet, int sheetBodyAddRowCount)
        {
            if (config.Foot.ConfigCellList == null || config.Foot.ConfigCellList.Count <= 0)
            {
                return;
            }

            var dictConfigSource = configSource.Foot.CellsInfoList.ToDictionary(a => a.ConfigValue);
            foreach (var item in config.Foot.ConfigCellList)
            {
                if (configSource.Foot == null ||
                    configSource.Foot.CellsInfoList == null ||
                    configSource.Foot.CellsInfoList.Count == 0) //excel中有配置foot,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
                {
                    break;
                }

                //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值
                var cellPoint = new ExcelCellPoint(item.Address);
                string colMapperName = item.ConfigValue;

                object val = config.Foot.ConfigItemMustExistInDataColumn
                    ? dictConfigSource[item.ConfigValue].FillValue
                    : dictConfigSource.ContainsKey(item.ConfigValue) ? dictConfigSource[item.ConfigValue].FillValue : null;

                ExcelRange cells = worksheet.Cells[cellPoint.Row + sheetBodyAddRowCount, cellPoint.Col];
                if (config.Foot.CellCustomSetValue != null)
                {
                    config.Foot.CellCustomSetValue.Invoke(colMapperName, val, cells);
                }
                else
                {
                    SetWorksheetCellsValue(config, cells, val, colMapperName); //item.Key -> D13 , item.Value -> 总计
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="cells">s结尾表示单元格有可能是合并单元格</param>
        /// <param name="val">值</param>
        /// <param name="colMapperName">excel填充的列名,不想传值请使用null,用来确保填充的数据格式,譬如身份证, 那么单元格必须要是</param> 
        private static void SetWorksheetCellsValue(EPPlusConfig config, ExcelRange cells, object val, string colMapperName)
        {
            cells.Value = config.UseFundamentals ? config.CellFormatDefault(colMapperName, val, cells) : val;
            //注:排除3种值( DBNull.Value , null , "") 后 如果 cells.Value 仍然没有值,有可能是配置的单元格以 '开头.
            //譬如: '$tb1Id. 对于这种配置我程序无法检测出来(或者说我没找到检测'开头的方法)
            //下面代码有问题,当遇到日期类型的时候, 值是赋值上去的,但是 cells.value 却!= val .所以下面代码注释
            //if (val != DBNull.Value  && val != null && val != "" && cells.Value != val)
            //{
            //    //如果值没赋值上去,那么抛异常
            //    throw new Exception($"工作簿'{cells.Worksheet.Name}'的配置列'{colMapperName}'的单元格格式有问题,程序无法将值'{val}'赋值到对应的单元 格'{cells.Address}'中.配置的单元格中可能是'开头的,请把'去掉");
            //}
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
                var isMergeCell = EPPlusHelper.IsMergeCell(args.ws, args.DataTitleRow, col, out var mergeCellAddress);
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

                var colName = EPPlusHelper.GetMergeCellText(args.ws, ea.Start.Row, ea.Start.Column);
                if (string.IsNullOrEmpty(colName)) break;
                colName = ExtractName(colName);
                if (string.IsNullOrEmpty(colName)) break;

                dataColEndActual = newDataColEndActual;

                if (args.POCO_Property_AutoRename_WhenRepeat)
                {
                    AutoRename(colNameList, nameRepeatCounter, colName, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
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

        /// <summary>
        /// 提取符合c#规范的名字
        /// </summary>
        /// <param name="colName"></param>
        /// <returns></returns>
        private static string ExtractName(string colName)
        {
            string reg = @"[_a-zA-Z\u4e00-\u9FFF][A-Za-z0-9_\u4e00-\u9FFF]*";//去掉不合理的属性命名的字符串(提取合法的字符并接成一个字符串)
            colName = RegexHelper.GetStringByReg(colName, reg).Aggregate("", (current, item) => current + item);
            return colName;
        }

        /// <summary>
        /// 设置报表(能折叠行的excel)格式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="config"></param>
        /// <param name="destRow"></param>
        /// <param name="maxIntervalRow"></param>
        private static void SetReport(ExcelWorksheet worksheet, DataRow row, EPPlusConfig config, int destRow, int maxIntervalRow = 0)
        {
            int level = Convert.ToInt32(row[config.Report.RowLevelColumnName]) - 1;//level是从0开始的
            for (int i = destRow; i <= destRow + maxIntervalRow; i++)//for循环主要用于多行合并的worksheet
            {
                worksheet.Row(i).OutlineLevel = level;
                worksheet.Row(i).Collapsed = config.Report.Collapsed;
            }
            //对于没有合并行的单元格来说完全可以用如下写法
            //worksheet.Row(destRow).OutlineLevel = level;
            //worksheet.Row(destRow).Collapsed = config.Report.Collapsed;
            worksheet.OutLineSummaryBelow = config.Report.OutLineSummaryBelow;
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
            if (isNullable_Boolean && (value == null || value.Length <= 0))
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
            if (isNullable_DateTime && (value == null || value.Length <= 0))
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
            if (isNullable_sbyte && (value == null || value.Length <= 0))
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
            if (isNullable_byte && (value == null || value.Length <= 0))
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
            if (isNullable_UInt16 && (value == null || value.Length <= 0))
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
            if (isNullable_Int16 && (value == null || value.Length <= 0))
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
            if (isNullable_UInt32 && (value == null || value.Length <= 0))
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
            if (isNullable_Int32 && (value == null || value.Length <= 0))
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
            if (isNullable_UInt64 && (value == null || value.Length <= 0))
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
            if (isNullable_Int64 && (value == null || value.Length <= 0))
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
            if (isNullable_float && (value == null || value.Length <= 0))
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
            if (isNullable_double && (value == null || value.Length <= 0))
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
            if (isNullable_decimal && (value == null || value.Length <= 0))
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
                if (value == null || value.Length <= 0)
                {
                    pInfo.SetValue(model, null);
                    return;
                }
                value = ExtractName(value);
                var enumType = pInfo_PropertyType.GetProperty("Value").PropertyType;
                TryThrowExceptionForEnum(pInfo, model, value, enumType, pInfo_PropertyType);
                var enumValue = Enum.Parse(enumType, value);
                pInfo.SetValue(model, enumValue);
                return;
            }
            if (pInfo_PropertyType.IsEnum)
            {
                if ((value == null || value.Length <= 0))
                {
                    throw new ArgumentException($@"无效的{pInfo_PropertyType.FullName}枚举值", pInfo.Name, new FormatException($"单元格值:{value}未被识别为有效的 {pInfo_PropertyType}(Enum类型)"));
                }
                value = ExtractName(value);
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
            if (attrs.Length <= 0)
            {
                return;
            }

            var attr = (EnumUndefinedAttribute)attrs[0];
            if (attr.Args == null || attr.Args.Length <= 0)
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
        private static string FormatAttributeMsg<T>(string pInfo_Name, T model, string value, string attrErrorMessage, string[] attrArgs) where T : class, new()
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
                    if (prop == null)
                    {
                        continue;
                    }

                    attrArgs[i] = prop.GetValue(model).ToString();
                }
            }

            string message = string.Format(attrErrorMessage, attrArgs);
            return message;
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

        public static EPPlusConfigSource GetEmptyConfigSource() => new EPPlusConfigSource()
        {
            Head = new EPPlusConfigSourceHead(),
            Body = new EPPlusConfigSourceBody(),
            Foot = new EPPlusConfigSourceFoot(),
        };

        #endregion

        #region 设置Head与foot配置的数据源

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
            //        dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
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
            //        dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
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

        public static T InitGetExcelListArgsModel<T>() where T : class
        {
            var t_ctor = typeof(T).GetConstructor(new Type[] { });
            if (t_ctor == null)
            {
                return null;
            }

            var model = t_ctor.Invoke(new object[] { });

            //foreach (var p in ReflectionHelper.GetProperties(typeof(T)).Where(p => p == typeof(KV<,>)))
            foreach (var p in ReflectionHelper.GetProperties(typeof(T)))
            {
                if (!p.PropertyType.HasImplementedRawGeneric(typeof(KV<,>)))
                {
                    continue;
                }

                var p_ctor = p.PropertyType.GetConstructor(new Type[] { });
                if (p_ctor == null)
                {
                    continue;
                }
                p.SetValue(model, p_ctor.Invoke(new object[] { }));
            }
            return (T)model;
        }

        public static GetExcelListArgs GetExcelListArgsDefault(ExcelWorksheet ws, int rowIndex)
        {
            var args = new GetExcelListArgs
            {
                ws = ws,
                DataRowStart = rowIndex,
                DataTitleRow = rowIndex - 1,
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
                KVSource = new KVSource(),
            };
            return args;
        }

        public static GetExcelListArgs<T> GetExcelListArgsDefault<T>(ExcelWorksheet ws, int rowIndex) where T : class
        {
            var argsReturn = new GetExcelListArgs<T>
            {
                HavingFilter = null,
                WhereFilter = null,
                Model = InitGetExcelListArgsModel<T>()
            };
            var args = GetExcelListArgsDefault(ws, rowIndex);
            var dict = ReflectionHelper.GetProperties(typeof(GetExcelListArgs)).ToDictionary(item => item.Name, item => item.GetValue(args));

            foreach (var item in ReflectionHelper.GetProperties(typeof(GetExcelListArgs<T>)))
            {
                if (dict.ContainsKey(item.Name))
                {
                    item.SetValue(argsReturn, dict[item.Name]);
                }
            }

            return argsReturn;
        }

        /// <summary>
        /// 只能是最普通的excel.(每个单元格都是未合并的,第一行是列名,数据从第一列开始填充的那种.)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowIndex">数据起始行(不含列名),从1开始</param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex) where T : class, new()
        {
            var args = EPPlusHelper.GetExcelListArgsDefault<T>(ws, rowIndex);
            return EPPlusHelper.GetList<T>(args);
        }

        /// <summary>
        /// 只能是最普通的excel.(第一行是必须是列名,数据填充列起始必须是A2单元格,且每个单元格都是未合并的)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowIndex">数据起始行(不含列名),从1开始</param> 
        /// <param name="everyCellReplaceOldValue"></param>
        /// <param name="everyCellReplaceNewValue"></param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellReplaceOldValue, string everyCellReplaceNewValue) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowIndex);
            if (everyCellReplaceOldValue != null && everyCellReplaceNewValue != null)
            {
                args.EveryCellReplaceList = new Dictionary<string, string> { { everyCellReplaceOldValue, everyCellReplaceNewValue } };
            }
            return GetList<T>(args);
        }

        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellPrefix, Dictionary<string, string> everyCellReplace) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowIndex);
            args.EveryCellPrefix = everyCellPrefix;
            args.EveryCellReplaceList = everyCellReplace;
            return GetList<T>(args);
        }

        private static bool DynamicCalcStep(ScanLine scanLine)
        {
            if (scanLine == ScanLine.SingleLine) return false;
            if (scanLine == ScanLine.MergeLine) return true; //在代码的while中进行动态计算
            throw new Exception("不支持的ScanLine");
        }

        public static List<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {
            var colNameList = GetExcelColumnOfModel(args);//主要是计算DataColEnd的值, 放在第一行还是因为 单元测试 03.02的示例

            void CheCk()
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
                        if (!EPPlusHelper.IsMergeCell(args.ws, row: rowNo, col: args.DataColStart, out var mergeCellAddress))
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

                            if (!EPPlusHelper.IsMergeCell(args.ws, row: rowNo, col: colNo))
                            {
                                throw new Exception($@"检测到数据的起始列是合并行,请确保当前行的数据都是合并行.当前{new ExcelCellPoint(rowNo, colNo).R1C1}单元格不满足需求."); //参考 示例 03.14  或03.02的模版
                            }
                        }
                    }
                }

                #endregion

            }

            CheCk();

            Type type = typeof(T);

            #region 获得字典
            var dictModelPropNameExistsExcelColumn = new Dictionary<string, bool>();//Model属性在Excel列中存在, key: ModelPropName
            var dictModelPropNameToExcelColumnName = new Dictionary<string, string>();//Model属性名字对应的excel的标题列名字
            var dictExcelColumnIndexToModelPropName_Temp = new Dictionary<int, string>();//Excel的列标题和Model属性名字的映射
                                                                                         //初始化上面3个dict
            foreach (var props in type.GetProperties())
            {
                dictModelPropNameExistsExcelColumn.Add(props.Name, false);
                dictModelPropNameToExcelColumnName.Add(props.Name, null);

                var propAttr_DisplayExcelColumnName = ReflectionHelper.GetAttributeForProperty<DisplayExcelColumnNameAttribute>(type, props.Name);
                if (propAttr_DisplayExcelColumnName.Length > 0)
                {
                    dictModelPropNameToExcelColumnName[props.Name] = ((DisplayExcelColumnNameAttribute)propAttr_DisplayExcelColumnName[0]).Name;
                }
                var propAttr_ExcelColumnIndex = ReflectionHelper.GetAttributeForProperty<ExcelColumnIndexAttribute>(type, props.Name);
                if (propAttr_ExcelColumnIndex.Length > 0)
                {
                    dictExcelColumnIndexToModelPropName_Temp.Add(((ExcelColumnIndexAttribute)propAttr_ExcelColumnIndex[0]).Index, props.Name);
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

                if (pInfo == null)
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

                        if ((_matchingModel & MatchingModel.eq) != MatchingModel.eq) continue;
                        if (dictMatchingModelException.ContainsKey(matchingModelValue)) continue;//如果已经添加过了

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
                        if ((_matchingModel & MatchingModel.gt) != MatchingModel.gt) continue;
                        if (dictMatchingModelException.ContainsKey(MatchingModel.gt)) continue;
                        dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, args.ws));
                    }
                    else if (matchingModelValue == MatchingModel.lt)
                    {
                        if ((_matchingModel & MatchingModel.lt) != MatchingModel.lt) continue;
                        if (dictMatchingModelException.ContainsKey(MatchingModel.lt)) continue;
                        dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, args.ws));
                    }
                    else if (matchingModelValue == MatchingModel.neq)
                    {
                        if ((_matchingModel & MatchingModel.neq) != MatchingModel.neq) continue;
                        //neq 会调用 gt+ lt ,所以要排除,即 _matchingModel的值 不能是带neq的标志枚举的值
                        if ((_matchingModel & MatchingModel.gt) == MatchingModel.gt) continue;
                        if ((_matchingModel & MatchingModel.lt) == MatchingModel.lt) continue;

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

            //var dictColName = colNameList.ToDictionary(item => new ExcelCellPoint(item.ExcelAddress).Col, item => item);// key是第n列

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList == null
                ? GetExcelListArgs.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;

            //var ctor = type.GetConstructor(new Type[] { });
            //if (ctor == null) throw new ArgumentException($"通过反射无法得到'{type.FullName}'的一个无构造参数的构造器.");

            var dictPropAttrs = new Dictionary<string, Dictionary<string, Attribute>>();//属性里包含的Attribute

            #region 内置的Attribute
            var dictUnique = new Dictionary<string, Dictionary<string, bool>>();//属性的 UniqueAttribute
            var dictKVSet = new Dictionary<string, Dictionary<string, bool>>();//属性的 KVSetAttribute
            string key_UniqueAttribute = typeof(UniqueAttribute).FullName;
            string key_KVSetAttribute = typeof(KVSetAttribute).FullName;

            var cache_PropertyInfo = new Dictionary<string, PropertyInfo>();
            foreach (var excelCellInfo in colNameList)
            {
                if (!GetPropName(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All, out var propName))
                {
                    continue;
                }

                var pInfo = GetPropertyInfo(cache_PropertyInfo, propName, type);

                #region 初始化Attr要处理相关的数据
                dictPropAttrs.Add(pInfo.Name, new Dictionary<string, Attribute>());//这里new 的Dict 的key 代表的是Attribute的FullName

                var uniqueAttrs = ReflectionHelper.GetAttributeForProperty<UniqueAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (uniqueAttrs.Length > 0)
                {
                    dictPropAttrs[pInfo.Name].Add(key_UniqueAttribute, (UniqueAttribute)uniqueAttrs[0]);
                    dictUnique.Add(pInfo.Name, new Dictionary<string, bool>());
                }

                var KVSetAttrs = ReflectionHelper.GetAttributeForProperty<KVSetAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (KVSetAttrs.Length > 0)
                {
                    if (args.KVSource == null || args.KVSource.Count <= 0)
                    {
                        throw new ArgumentException($@"检测到KVSetAttribute,但是KVSource却未配置");
                    }
                    dictPropAttrs[pInfo.Name].Add(key_KVSetAttribute, (KVSetAttribute)KVSetAttrs[0]);
                    dictKVSet.Add(pInfo.Name, new Dictionary<string, bool>());

                }

                #endregion
            }

            #endregion

            #region 获得 list

            ExcelCellInfoNeedTo(args.ReadCellValueOption, out var toTrim, out var toMergeLine, out var toDBC);

            var allException = args.GetList_NeedAllException ? new List<Exception>() : null;

#if DEBUG
            var debugvar_whileCount = 0;
#endif
            Func<object[], object> deletgateCreateInstance = ExpressionTreeExtensions.BuildDeletgateCreateInstance(type, new Type[0]);

            List<T> list = new List<T>();
            var dynamicCalcStep = DynamicCalcStep(args.ScanLine);
            Exception exception = null;
            int row = args.DataRowStart;

            while (true)//遍历行, 异常或者出现空行,触发break;
            {
                if (args.ws.Row(row).Hidden)
                {
                    throw new Exception($@"工作簿:'{args.ws.Name}'不允许存在隐藏行,检测到第{row}行是隐藏行");
                }
#if DEBUG
                debugvar_whileCount++;
#endif
                //判断整行数据是否都没有数据
                bool isNoDataAllColumn = true;

                //Sample02_3,12000的数据
                //T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转  1.2-2.1秒
                //T model = type.CreateInstance<T>();//3秒+
                T model = (T)deletgateCreateInstance(null); //上面的方法给拆开来 . 1.1-1.4

                foreach (var excelCellInfo in colNameList)
                {
                    if (!GetPropName(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All, out var propName))
                    {
                        continue;
                    }

                    var pInfo = GetPropertyInfo(cache_PropertyInfo, propName, type);
                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

#if DEBUG
                    string value;
                    if (pInfo.PropertyType == typeof(DateTime?) || pInfo.PropertyType == typeof(DateTime))
                    {
                        //todo:对于日期类型的,有时候要获取Cell.Value, 有空了修改
                        value = GetMergeCellText(args.ws, row, col);
                    }
                    else
                    {
                        value = GetMergeCellText(args.ws, row, col);
                    }
#else
                    string value =  GetMegerCellText(ws, row, col);
#endif

                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);

                    var propAttrs = dictPropAttrs[pInfo.Name];//当前属性的所有特性]

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

                        //var propAttrs = dictPropAttrs[pInfo.Name];//当前属性的所有特性

                        if (propAttrs.ContainsKey(key_UniqueAttribute))
                        {
                            #region uniqueAttr的具体实现
                            var uniqueAttr = (UniqueAttribute)propAttrs[key_UniqueAttribute];

                            if (!valueIsNullOrEmpty)
                            {
                                if (!dictUnique[pInfo.Name].ContainsKey(value))
                                {
                                    dictUnique[pInfo.Name].Add(value, default(bool));
                                }
                                else
                                {
                                    string exception_msg = string.IsNullOrEmpty(uniqueAttr.ErrorMessage) ? $@"属性'{pInfo.Name}'的值:'{value}'出现了重复" : uniqueAttr.ErrorMessage;
                                    throw new ArgumentException(exception_msg, pInfo.Name);
                                }
                            }
                            #endregion
                        }
                        if (propAttrs.ContainsKey(key_KVSetAttribute))
                        {
                            var kvsetAttr = (KVSetAttribute)propAttrs[key_KVSetAttribute];
                            var haveKvsource = args.KVSource.ContainsKey(kvsetAttr.Name);
                            if (kvsetAttr.MustInSet && !haveKvsource)
                            {
                                if (!string.IsNullOrEmpty(kvsetAttr.ErrorMessage) && kvsetAttr.ErrorMessage.Length > 0)
                                {
                                    var msg = FormatAttributeMsg(pInfo.Name, model, value, kvsetAttr.ErrorMessage, kvsetAttr.Args);
                                    throw new ArgumentException(msg);
                                }
                                else
                                {
                                    throw new ArgumentException($@"属性'{pInfo.Name}'的值:'{value}'未找到对应的集合列表", pInfo.Name);
                                }
                            }

                            object kvsource = args.KVSource[kvsetAttr.Name];
                            var kvsourceType = kvsource.GetType();

                            //var is_kvsourceType = kvsourceType.GetGenericTypeDefinition() == typeof(KVSource<,>);
                            var is_kvsourceType = kvsourceType.HasImplementedRawGeneric(typeof(KvSource<,>));

                            if (is_kvsourceType)
                            {
                                var kvsourceTypeTKey = kvsourceType.GenericTypeArguments[0];
                                //var kvsourceTypeTValue = kvsourceType.GenericTypeArguments[1];

                                var prop_kvsource = (IKVSource)kvsource;

                                prop_kvsource.GetInfoByKey(value, out bool kv_Value_inKvSource, out object kv_Value,
                                    out bool haveState, out object state);

                                if (!kv_Value_inKvSource && kvsetAttr.MustInSet)
                                {
                                    var msg = string.IsNullOrEmpty(kvsetAttr.ErrorMessage)
                                        ? $@"属性'{pInfo.Name}'值:'{value}'未在'{kvsetAttr.Name}'集合中出现."
                                        : FormatAttributeMsg(pInfo.Name, model, value, kvsetAttr.ErrorMessage, kvsetAttr.Args);// string.Format(kvsetAttr.ErrorMessage, kvsetAttr.Args);
                                    throw new ArgumentException(msg, pInfo.Name);
                                }

                                var typeKVArgs = pInfo.PropertyType.GetGenericArguments();
                                var typeKV = typeof(KV<,>).MakeGenericType(typeKVArgs);

                                object[] invokeConstructorParameters;
                                //if (typeKVArgs[0].FullName == typeof(string).FullName)
                                //{
                                //    invokeParameters = new object[] { value, kv_Value };
                                //}
                                //else
                                //{
                                //    invokeParameters = new object[] { Convert.ChangeType(value, typeKVArgs[0]), kv_Value };
                                //}
                                //代码和上面的是一样的效果,这个更加方便阅读
                                if (kvsourceTypeTKey == typeof(string))
                                {

                                    invokeConstructorParameters = new object[] { value, kv_Value };
                                }
                                else
                                {
                                    invokeConstructorParameters = new object[] { Convert.ChangeType(value, kvsourceTypeTKey), kv_Value };
                                }

                                var modelValue = typeKV.GetConstructor(typeKVArgs).Invoke(invokeConstructorParameters);

                                if (kv_Value == null) //上面Invoke时, 是调用2个参数的构造方法的,所以,这里要修正HasValue值
                                {
                                    if (!kv_Value_inKvSource)//因为默认值是true,所以,只要修改值为false的情况就可以了
                                    {
                                        typeKV.GetProperty("HasValue").SetValue(modelValue, false);
                                    }
                                }
                                if (haveState)
                                {
                                    typeKV.GetProperty("HasState").SetValue(modelValue, true);
                                    typeKV.GetField("_state", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(modelValue, state);
                                }

                                pInfo.SetValue(model, modelValue);
                            }
                        }
                        #endregion
                    }
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

                        //赋值, 注:遇到 KV<,> 类型的统一不处理
                        if (!pInfo.PropertyType.HasImplementedRawGeneric(typeof(KV<,>)))
                        {
                            GetList_SetModelValue(pInfo, model, value);
                        }
                    }
                    catch (ArgumentException e)
                    {
                        exception = new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}", e);
                        break;
                    }
                    catch (ValidationException e)
                    {
                        exception = new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}({pInfo.Name}:{e.Message})", e);
                        //log $"无效的单元格:{new ExcelCellAddress(row, col).Address},'{model.GetType().FullName}'类型的'{pInfo.Name}'属性验证未通过:'{e.Message}'"
                        break;
                    }
                    catch (Exception e)
                    {
                        exception = e;
                        break;
                    }
                }

                if (isNoDataAllColumn)
                {
                    if (row == args.DataRowStart)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    var isEmptyLine = true;

                    #region 修改 isEmptyLine  代码来自 lable1:遍历属性

                    //上面代码(lable1:遍历属性) 对 isNoDataAllColumn 的判断逻辑有问题. 即:Excel第一列(Sequence)对应的Model是Int类型, 且Sequence忘记填写时,变量isNoDataAllColumn的值则是错误的.

                    foreach (var excelCellInfo in colNameList)
                    {
                        if (!GetPropName(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All, out var propName))
                        {
                            continue;
                        }

                        var pInfo = GetPropertyInfo(cache_PropertyInfo, propName, type);
                        var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

#if DEBUG
                        string value;
                        if (pInfo.PropertyType == typeof(DateTime?) || pInfo.PropertyType == typeof(DateTime))
                        {
                            //todo:对于日期类型的,有时候要获取Cell.Value, 有空了修改
                            value = GetMergeCellText(args.ws, row, col);
                        }
                        else
                        {
                            value = GetMergeCellText(args.ws, row, col);
                        }
#else
                        string value =  GetMegerCellText(ws, row, col);
#endif

                        if (value.Length > 0)
                        {
                            isEmptyLine = false;
                            break;
                        }
                    }

                    #endregion

                    if (isEmptyLine)
                    {
                        break; //出现空行,读取模版结束
                    }
                }

                //1.添加Step,准备读取下一行数据
                if (dynamicCalcStep)
                {
                    //while里面动态计算
                    if (EPPlusHelper.IsMergeCell(args.ws, row, col: args.DataColStart, out var mergeCellAddress))
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

                //2.如果有异常,抛出异常,不对下一行进行读取
                if (exception != null)
                {
                    if (args.GetList_NeedAllException)
                    {
                        allException.Add(exception);
                        exception = null;
                        continue;
                    }
                    else
                    {
                        throw exception;
                    }
                }

                if (args.WhereFilter == null || args.WhereFilter.Invoke(model))
                {
                    list.Add(model);
                }

            }

#if DEBUG
            Console.WriteLine("DEBUG only --- 变量debugvar_whileCount的值:" + debugvar_whileCount);
#endif

            var keyWithExceptionMessageStart = "无效的单元格:";
            if (allException != null && allException.Count > 0)
            {
                bool allExceptionIsArgumentException = true;
                var errGroupMsg = new Dictionary<string, List<string>>();

                foreach (var ex in allException)
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
                    throw new AggregateException(allException);
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

                    sb2.RemoveLastChar(',');
                    if (sb2.Length > 0)
                    {
                        sb.Append($"({sb2}),");
                    }
                }

                var argumentExceptionMsg = sb.RemoveLastChar(',').ToString();
                throw new ArgumentException(argumentExceptionMsg);
            }

            #endregion

            return args.HavingFilter == null ? list : list.Where(item => args.HavingFilter.Invoke(item)).ToList();
        }

        private static void ExcelCellInfoNeedTo(ReadCellValueOption readCellValueOption, out bool toTrim, out bool toMergeLine,
            out bool toDBC)
        {
            toTrim = (readCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            toMergeLine = (readCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            toDBC = (readCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;
        }

        private static PropertyInfo GetPropertyInfo(Dictionary<string, PropertyInfo> cache_PropertyInfo, string propName, Type type)
        {
            if (!cache_PropertyInfo.ContainsKey(propName))
            {
                var pInfo2 = type.GetProperty(propName);
                if (pInfo2 == null) //防御式编程判断
                {
                    throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                }

                cache_PropertyInfo.Add(propName, pInfo2);
            }

            PropertyInfo pInfo = cache_PropertyInfo[propName];
            return pInfo;
        }

        /// <summary>
        /// 获得属性名
        /// </summary>
        /// <param name="ExcelAddress"></param>
        /// <param name="dictExcelAddressCol"></param>
        /// <param name="dictExcelColumnIndexToModelPropName_All"></param>
        /// <param name="propName"></param>
        /// <returns>是否get到了PropName</returns>
        private static bool GetPropName(ExcelAddress ExcelAddress, Dictionary<ExcelAddress, int> dictExcelAddressCol,
            Dictionary<int, string> dictExcelColumnIndexToModelPropName_All, out string propName)
        {
            int excelCellInfo_ColIndex = dictExcelAddressCol[ExcelAddress];
            if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null) //不存在,跳过
            {
                propName = null;
                return false;
            }
            propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
            return !string.IsNullOrEmpty(propName);
        }

        public static DataTable GetDataTable(GetExcelListArgs<DataRow> args)
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.DataRowStart;
            if (rowIndex <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowIndex}'错误,值应该大于0");
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

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList == null
                ? GetExcelListArgs.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;


            #region 获得 list

            int row = rowIndex;
            bool dynamicCalcStep = DynamicCalcStep(args.ScanLine);
            ExcelCellInfoNeedTo(args.ReadCellValueOption, out var toTrim, out var toMergeLine, out var toDBC);
            while (true)
            {
                bool isNoDataAllColumn = true;//判断整行数据是否都没有数据
                var dr = dt.NewRow();

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    string propName = excelCellInfo.Value.ToString();

                    if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

                    string value = GetMergeCellText(ws, row, col);
                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);

                    if (!valueIsNullOrEmpty)
                    {
                        isNoDataAllColumn = false;

                        #region 判断每个单元格的开头
                        if (args.EveryCellPrefix?.Length > 0)
                        {
                            var indexof = value.IndexOf(args.EveryCellPrefix);
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
                    if (row == rowIndex)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    break; //出现空行,读取模版结束
                }
                //else
                if (args.WhereFilter == null || args.WhereFilter.Invoke(dr))
                {
                    dt.Rows.Add(dr);
                }
                if (dynamicCalcStep)
                {
                    if (EPPlusHelper.IsMergeCell(ws, row, col: 1, out var mergeCellAddress))
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

            return args.HavingFilter == null ? dt : dt.AsEnumerable().Where(item => args.HavingFilter.Invoke(item)).CopyToDataTable();
        }

        private static string DealMatchingModelException(MatchingModelException matchingModelException)
        {
            if ((matchingModelException.MatchingModel & MatchingModel.eq) == MatchingModel.eq)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
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

                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
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
                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
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
            if (dictExcelColumnIndexToModelPropNameAll == null) throw new ArgumentNullException(nameof(dictExcelColumnIndexToModelPropNameAll));
            if (dictModelPropNameExistsExcelColumn == null) throw new ArgumentNullException(nameof(dictModelPropNameExistsExcelColumn));

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
                if (dictExcelColumnIndexToModelPropNameAll[excelColumnIndex] == null)
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

        #region 读取excel配置

        /// <summary>
        /// 从Excel中获取默认的配置信息
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetIndex">第几个workSheet作为模版,从1开始</param>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, int workSheetIndex)
        {
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
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
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));
            var worksheet = GetExcelWorksheet(excelPackage, workSheetName);
            EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);
            SetConfigBodyFromExcel_OtherPara(config, worksheet);
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

                var mergedCellsList = worksheet.MergedCells.ToList();
                foreach (var configCellInfo in config.Body[nth].Option.ConfigLine)
                {
                    if (worksheet.Cells[configCellInfo.Address].Merge) //item.Address  D4
                    {
                        var addressPrecise = EPPlusHelper.GetMergeCellAddressPrecise(worksheet, configCellInfo.Address); //D4:E4格式的
                        allConfigInterval += new ExcelCellRange(addressPrecise).IntervalCol;

                        configCellInfo.FullAddress = addressPrecise;
                        configCellInfo.IsMergeCell = true;

                    }
                    else
                    {
                        var mergeCellAddress = mergedCellsList.Find(a => a.Contains(configCellInfo.Address));
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
            //遇到问题描述:创建一个excel,在C7,C8,C9,10单元格写入一些字符串, sheet.Cells.Value 是object[4,3]的数组, 但我要的是object[10,3]的数组
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value == null)
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
            if (arr == null)
            {
                throw new ArgumentNullException(nameof(arr));
            }

            var returnType = typeof(TOut);
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;
                    if (arr[i, j].ToString().Length <= 0) continue;
                    if (returnType == typeof(ExcelCellPoint))
                    {
                        var cell = new ExcelCellPoint(i + 1, j + 1);
                        return cell;
                    }
                    if (returnType == typeof(ExcelCellRange))
                    {
                        var mergeCellAddress = GetMergeCellAddressPrecise(ws, i + 1, j + 1);
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
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (cellStr.Length < 3) //配置至少有4个字符.所以,4个以下的都可以跳过
                    {
                        continue; //不用""比较, .Length速度比较快
                    }

                    //var cell = sheet.Cells[i + 1, j + 1];//当单元格值是公式时,没法在configLine里进行add, 因为下面的 nthStr 是 ""

                    if (!cellStr.StartsWith("$tb")) continue;

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
                            if (newKey == null)
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
                foreach (var item in bodyConfig.Value.Option.ConfigExtra.GetRepeatBy(a => new { a.ConfigValue }))
                {
                    sb.Append($@"{item.Address}-{item.ConfigValue},");
                }
                if (sb.RemoveLastChar(',').Length > 0)
                {
                    throw new ArgumentException($"Excel文件中的$tbs{bodyConfig.Key}部分配置了相同的项:{sb}");
                }

                sb.Clear();
                foreach (var item in bodyConfig.Value.Option.ConfigLine.GetRepeatBy(a => new { a.ConfigValue }))
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
            if (!startWith.StartsWith("$")) throw new ArgumentException("配置项必须是$开头");

            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");

            var fixedCellsInfoList = new List<EPPlusConfigFixedCell>();
            var replaceStr = startWith.RemovePrefix("$");
            for (var i = 0; i < arr.GetLength(0); i++)
            {
                for (var j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (!cellStr.StartsWith(startWith)) continue;

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

        #region 一些帮助方法

        /// <summary>
        /// 将workSheetIndex转换为代码中确切的值
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始</param>
        /// <returns></returns>
        private static int ConvertWorkSheetIndex(ExcelPackage excelPackage, int workSheetIndex)
        {
            //if (!excelPackage.Compatibility.IsWorksheets1Based)
            //{
            //    workSheetIndex -= 1; //从0开始的, 需要自己 -1;
            //}
            //return workSheetIndex;

            //var offset = excelPackage.Compatibility.IsWorksheets1Based ? 0 : -1;
            //return workSheetIndex + offset;

            return workSheetIndex + (excelPackage.Compatibility.IsWorksheets1Based ? 0 : -1);
        }

        /// <summary>
        /// 获得精确的合并单元格地址
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, int row, int col)
        {
            return EPPlusHelper.IsMergeCell(ws, row, col, out var mergeCellAddress)
                ? mergeCellAddress
                : new ExcelCellPoint(row, col).R1C1;
        }

        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, string r1c1)
        {
            var excelRange = new ExcelCellRange(r1c1);
            if (excelRange.End.Col == 0) //r1c1 为 D4  这种值
            {
                return GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.Start.Col);
            }
            else
            {
                return GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.End.Col);
            }
        }

        public static string GetLeftCellAddress(ExcelWorksheet ws, string currentCellAddress)
        {
            var ea = new ExcelAddress(currentCellAddress);
            var row = ea.Start.Row;
            var col = ea.Start.Column;

            if (!EPPlusHelper.IsMergeCell(ws, row: row, col: col, out var mergeCellAddress))
            {
                return new ExcelCellPoint(row, col - 1).R1C1;
            }

            var mergeCell = new ExcelAddress(mergeCellAddress);
            var leftCellRow = mergeCell.Start.Row;
            var leftCellCol = mergeCell.Start.Column - 1;
            if (EPPlusHelper.IsMergeCell(ws, row: leftCellRow, col: leftCellCol, out var leftCellAddress))
            {
                return new ExcelAddress(leftCellAddress).Address;
            }
            else
            {
                return new ExcelCellPoint(leftCellRow, leftCellCol).R1C1; //左边的单元格是普通的单元格
            }
        }

        public static bool IsMergeCell(ExcelWorksheet ws, int row, int col)
        {
            return IsMergeCell(ws, row, col, out var mergeCellAddress);
        }

        /// <summary>
        /// 是不是合并单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="mergeCellAddress"></param>
        /// <returns></returns>
        public static bool IsMergeCell(ExcelWorksheet ws, int row, int col, out string mergeCellAddress)
        {
            mergeCellAddress = ws.MergedCells[row, col];
            return mergeCellAddress != null;
        }

        /// <summary>
        /// 获得合并单元格的值,如果不是合并单元格, 返回单元格的值
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellText(ExcelWorksheet ws, int row, int col)
        {
            var isMergeCell = EPPlusHelper.IsMergeCell(ws, row, col, out var mergeCellAddress);
            if (!isMergeCell) return GetCellText(ws, row, col);
            var ea = new ExcelAddress(mergeCellAddress);
            return ws.Cells[ea.Start.Row, ea.Start.Column].Text;
        }

        /// <summary>
        /// 如果是合并单元格,请传入第一个合并单元格的坐标
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">从1开始</param>
        /// <param name="col">从1开始</param>
        /// <param name="when_TextProperty_NullReferenceException_Use_ValueProperty"></param>
        /// <returns></returns>
        public static string GetCellText(ExcelWorksheet ws, int row, int col, bool when_TextProperty_NullReferenceException_Use_ValueProperty = true)
        {
            return GetCellText(ws.Cells[row, col], when_TextProperty_NullReferenceException_Use_ValueProperty);
        }

        public static string GetCellText(ExcelRange cell, bool when_TextProperty_NullReferenceException_Use_ValueProperty = true)
        {
            //if (cell.Merge) throw new Exception("没遇到过这个情况的");
            //return cell.Text; //这个没有科学计数法  注:Text是Excel显示的值,Value是实际值.
            try
            {
                //if (cell.Formula?.Length > 0)//cell 是公式
                //{

                //}
                return cell.Text;//有的单元格通过cell.Text取值会发生异常,但cell.Value却是有值的
            }
            catch (NullReferenceException)
            {
                if (when_TextProperty_NullReferenceException_Use_ValueProperty)
                {
                    return Convert.ToString(cell.Value);
                }
                throw;
            }
        }

        /// <summary>
        /// 根据dict检查指定单元格的值是否符合预先定义.
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="dict">k:r1c1, v:具体值</param>
        public static bool CheckWorkSheetCellValue(ExcelWorksheet ws, Dictionary<string, string> dict)
        {
            foreach (var key in dict.Keys)
            {
                var cell = new ExcelCellPoint(key);
                if (EPPlusHelper.GetCellText(ws, cell.Row, cell.Col) != dict[key])
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 冻结窗口面板
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">冻结[1,Row)行</param>
        /// <param name="column">冻结{1,Column) 列</param>
        public static void FreezePanes(ExcelWorksheet ws, int row, int column)
        {
            ws.View.FreezePanes(row, column);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileOutDirectoryName"></param>
        /// <param name="dataConfigInfo"></param>
        /// <param name="cellCustom"></param>
        /// <returns></returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(string filePath, string fileOutDirectoryName, List<ExcelDataConfigInfo> dataConfigInfo, Action<ExcelRange> cellCustom = null)
        {
            using (var ms = new MemoryStream())
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var defaultConfigList = FillExcelDefaultConfig(excelPackage, dataConfigInfo, cellCustom);

                var haveConfig = defaultConfigList.Find(a => a.ClassPropertyList.Count > 0) != null;
                if (haveConfig)
                {
                    excelPackage.SaveAs(ms);
                    ms.Position = 0;
                    ms.Save($@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result.xlsx");
                }

                return defaultConfigList;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="dataConfigInfo">指定的worksheet</param>
        /// <param name="cellCustom"></param>
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
                if (dataConfigInfo == null)
                {
                    list.Add(FillExcelDefaultConfig(ws, titleLine, titleColumn, cellCustom));
                    continue;
                }

                var configInfo = dataConfigInfo.Find(a => a.WorkSheetName == ws.Name);
                if (configInfo == null)
                {
                    continue;
                }
                var address = GetMergeCellAddressPrecise(ws, row: configInfo.TitleLine, col: configInfo.TitleColumn);
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
            if (dataConfigInfo == null)
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

        public static List<string> KeysTypeOfDecimal => new List<string>
        {
            "金额", "钱", "数额",
            "money", "Money", "MONEY",
            "amount", "Amount", "AMOUNT",
        };

        public static List<string> KeysTypeOfDateTime => new List<string>
        {
            "时间", "日期", "date", "Date", "DATE", "time", "Time", "TIME",
            "今天", "昨天", "明天", "前天",
            "day", "Day", "DAY",
            "tomorrow","Tomorrow","TOMORROW",
        };

        public static List<string> KeysTypeOfString => new List<string>
        {
            "序号", "编号", "id", "Id", "ID", "number", "Number", "NUMBER", "No",
            "身份证", "银行卡", "卡号", "手机", "座机",
            "mobile", "Mobile", "MOBILE", "tel", "Tel", "TEL",
        };

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
                var excelColName = ws.Cells[titleLineNumber, col].Merge ? GetMergeCellText(ws, titleLineNumber, col) : GetCellText(ws, titleLineNumber, col);

                var destColVal = ExtractName(excelColName).Trim().MergeLines();
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

                AutoRename(colNameList, nameRepeatCounter, thisColNameInfo, true);

                if (EPPlusHelper.IsMergeCell(ws, titleLineNumber, col, out mergeCellAddress))
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
            if (EPPlusHelper.IsMergeCell(ws, titleLineNumber, 1, out mergeCellAddress))
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
                ws.Cells[fillBodyLine, colNameList[i].ExcelColNameIndex].Value = $@"$tb1{(colNameList[i].IsRename ? colNameList[i].NameNew : colNameList[i].Name)}";
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

                if (KeysTypeOfDateTime.Any(item => propName.Contains(item)))
                {
                    sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(DateTime);");
                    sb_CreateClassSnippet.AppendLine($" public DateTime {propName} {{ get; set; }}");
                    sb_CrateClassSnippe_AppendLine_InForeach = true;
                }

                if (KeysTypeOfString.Any(item => propName.Contains(item)))
                {
                    sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(string);");
                    sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                    sb_CrateClassSnippe_AppendLine_InForeach = true;
                }

                if (KeysTypeOfDecimal.Any(item => propName.Contains(item)))
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

        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        private static void AutoRename(List<string> nameList, Dictionary<string, int> nameRepeatCounter, string name, bool renameFirstNameWhenRepeat)
        {

            if (!nameRepeatCounter.ContainsKey(name))
            {
                nameRepeatCounter.Add(name, 0);
            }

            if (!nameList.Contains(name) && nameRepeatCounter[name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name] == 1)
                {
                    for (int i = 0; i < nameList.Count; i++)
                    {
                        if (nameList[i] == name)
                        {
                            nameList[i] = nameList[i] + "1";
                            break;
                        }
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name];
                nameList.Add($@"{name}{++currentCounterVal}");
            }

            nameRepeatCounter[name] = ++nameRepeatCounter[name];
        }

        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        private static void AutoRename(List<ExcelCellInfoValue> nameList, Dictionary<string, int> nameRepeatCounter, ExcelCellInfoValue name, bool renameFirstNameWhenRepeat)
        {
            if (!nameRepeatCounter.ContainsKey(name.Name))
            {
                nameRepeatCounter.Add(name.Name, 0);
            }

            if (nameList.Find(a => a.Name == name.Name) == null && nameRepeatCounter[name.Name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name.Name] == 1)
                {
                    foreach (var t in nameList)
                    {
                        if (t.Name != name.Name) continue;
                        t.IsRename = true;
                        t.NameNew = t.Name + "1";
                        break;
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name.Name];
                name.IsRename = true;
                name.NameNew = $@"{name.Name}{++currentCounterVal}";
                nameList.Add(name);
            }
            nameRepeatCounter[name.Name] = ++nameRepeatCounter[name.Name];
        }

        /// <summary>
        /// 获得excel填写的配置内容
        /// </summary>
        /// <param name="content"></param>
        /// <param name="outResultPrefix"></param>
        /// <returns></returns>
        public static string GetFillDefaultConfig(string content, string outResultPrefix = "$tb1")
        {
            if (string.IsNullOrEmpty(content)) return content;
            content = content.TrimEnd();
            content.RemoveLastChar('\n');//excel选择列复制出来到文本上有换行,最后一个字符的ascii 是10 \n
            content.RemoveLastChar('\r');//如果是自己敲入的回车,那么也去掉
            var excel_cell_split = new char[] { '	', ' ', };// 两个单元格之间间隔的符号(\t),空格
            string[] splits = content.Split(excel_cell_split, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();
            StringBuilder sbColumn = new StringBuilder();
            foreach (var item in splits)
            {
                var newName = ExtractName(item);
                sb.Append($@"{outResultPrefix}{newName}{excel_cell_split[0]}");
                sbColumn.AppendLine($"dt.Columns.Add(\"{newName}\");");
            }

            sb.RemoveLastChar(excel_cell_split[0]);

            //sb.AppendLine().AppendLine().AppendLine();
            //sb.AppendLine($@"DataTable dt = new DataTable();");
            //sb.Append(sbColumn.ToString());

            return sb.ToString();
        }

        #region 获得单元格

        /// <summary>
        /// 根据值获的excel中对应的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static List<ExcelCellInfo> GetCellsBy(ExcelWorksheet ws, string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            return GetCellsBy(ws, ws.Cells.Value as object[,], a => a != null && a.ToString() == value);
        }

        /// <summary>
        /// 根据值获的excel中对应的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="cellsValue">一般通过ws.Cells.Value as object[,] 获得 </param>
        /// <param name="match">示例: a => a != null &amp;&amp; a.GetType() == typeof(string) &amp;&amp; ((string) a == "备注")</param>
        /// <returns></returns>
        public static List<ExcelCellInfo> GetCellsBy(ExcelWorksheet ws, object[,] cellsValue, Predicate<object> match)
        {
            if (cellsValue == null) throw new ArgumentNullException(nameof(cellsValue));

            var result = new List<ExcelCellInfo>();
            for (int i = 0; i < cellsValue.GetLength(0); i++)
            {
                for (int j = 0; j < cellsValue.GetLength(1); j++)
                {
                    if (match != null && match.Invoke(cellsValue[i, j]))
                    {
                        result.Add(new ExcelCellInfo
                        {
                            WorkSheet = ws,
                            ExcelAddress = new ExcelAddress(i + 1, j + 1, i + 1, j + 1),
                            Value = cellsValue[i, j]
                        });
                    }
                }
            }

            return result;
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

            using (var ms = new MemoryStream())
            {
                excelPackage.SaveAs(ms); // 导入数据到流中 
                ms.Position = 0;
                File.Delete(savePath); //删除文件。如果文件不存在,也不报错
                ms.Save(savePath);
            }

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
            StringBuilder sb = new StringBuilder("程序报错:");
            if (e.Message != null && e.Message.Length > 0)
            {
                sb.Append($@"Message:{e.Message}");
            }
            if (e.InnerException != null && e.InnerException.Message != null && e.InnerException.Message.Length > 0)
            {
                sb.Append($@"InnerExceptionMessage:{e.InnerException.Message}");
            }
            var txt = sb.ToString();
            return txt;
        }
        #endregion
    }
}
