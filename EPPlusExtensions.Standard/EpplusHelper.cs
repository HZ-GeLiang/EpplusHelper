using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using EPPlusExtensions.Attributes;
using EPPlusExtensions.Exceptions;
using EPPlusExtensions.Helper;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    public partial class EPPlusHelper
    {
        /// <summary>
        /// 填充Excel时创建的工作簿名字
        /// </summary>
        public static List<string> FillDataWorkSheetNameList = new List<string>();

        //类型参考网址: http://filext.com/faq/office_mime_types.php
        public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        #region GetExcelWorksheet

        /// <summary>
        /// 获得Excel的第N个Sheet
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始</param>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, int workSheetIndex)
        {
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));

            workSheetIndex = ConvertwsIndex(excelPackage, workSheetIndex);

            int sheetCount = excelPackage.Workbook.Worksheets.Count;
            if (workSheetIndex > sheetCount)
            {
                throw new ArgumentException($@"形参{nameof(workSheetIndex)}大于当前Excel的工作簿数量", nameof(workSheetIndex));//指定的参数已超出有效值的范围
            }
            return excelPackage.Workbook.Worksheets[workSheetIndex];
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
            var violateChars = new char[] { ':', '：', '\\', '/', '？', '*', '[', ']' };
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

        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, string workName, bool autoMappingWs)
        {
            if (workName == null) throw new ArgumentNullException(nameof(workName));
            var ws = excelPackage.Workbook.Worksheets[workName];
            if (autoMappingWs && ws == null && excelPackage.Workbook.Worksheets.Count == 1)
            {
                ws = excelPackage.Workbook.Worksheets[1];
            }
            if (ws == null) throw new ArgumentException($@"当前Excel中不存在名为'{workName}'的worksheet", nameof(workName));
            return ws;
        }
        /// <summary>
        /// 获得当前Excel的所有工作簿的名字
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        public static List<string> GetExcelWorksheetNames(ExcelPackage excelPackage)
        {
            List<string> wsNameList = new List<string>();
            for (int i = 1; i <= excelPackage.Workbook.Worksheets.Count; i++)
            {
                wsNameList.Add(GetExcelWorksheet(excelPackage, i).Name);
            }
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
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));

            if (excelPackage.Workbook.Worksheets.Count <= 1) //The workbook must contain at least one worksheet
            {
                return;
            }

            workSheetIndex = ConvertwsIndex(excelPackage, workSheetIndex);

            if (excelPackage.Workbook.Worksheets[workSheetIndex] != null)
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
            if (eWorkSheetHiddens == null) return wsNames;

            var count = excelPackage.Workbook.Worksheets.Count;
            for (int i = 1; i <= count; i++)
            {
                var ws = excelPackage.Workbook.Worksheets[i];
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

                //var cellpoint = new ExcelCellPoint(item.Key);
                string colMapperName = item.ConfigValue;
                object val = dictConfigSourceHead[item.ConfigValue].FillValue;
                //ExcelRange cells = worksheet.Cells[cellpoint.Row , cellpoint.Col];
                ExcelRange cells = worksheet.Cells[item.Address];

                if (config.Head.CellCustomSetValue != null)
                {
                    config.Head.CellCustomSetValue.Invoke(colMapperName, val, cells);
                }
                else
                {
                    //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值
                    //worksheet.Cells[item.Key].Value = configSource.SheetHead[item.Value];
                    //item.Key -> D2 item.Value -> 年龄
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
        /// <returns>新增了多行</returns>
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
                        var r1c1 = cellConfigInfo.Address;
                        int driftVale = 1; //浮动值,如果是合并单元格,则取合并单元格的行数
                        int delRow; //要删除的行号
                        if (r1c1.Contains(":")) //如果是合并单元格,修改浮动的行数
                        {
                            var cells = r1c1.Split(new[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                            if (cells.Length != 2) throw new Exception("该合并单元格的标识有问题,不是类似于A1:A2这个格式的");
                            int mergeCellStartRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[0]));
                            int mergeCellEndRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[1]));

                            driftVale = mergeCellEndRow - mergeCellStartRow + 1;
                            if (driftVale <= 0) throw new Exception("合并单元格的合并行数小于1");

                            delRow = mergeCellStartRow + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }
                        else //不是合并单元格
                        {
                            delRow = Convert.ToInt32(RegexHelper.GetLastNumber(r1c1)) + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
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
                var customValue = new CustomValue()
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
                            var val = row[colMapperName];
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

                            ExcelRange cells = worksheet.Cells[destRow, destStartCol, destRow + maxIntervalRow, destEndCol];

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

                    //第一遍循环:计算要插入多少行
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

                    if (insertRows > 0 && insertRowFrom > 0)
                    {
                        //在  InsertRowFrom 行前面插入 InsertRowCount 行.
                        //注:
                        //1. 新增的行的Height的默认值,需要自己修改
                        //2. copyStylesFromRow 的行计算是在 InsertRowFrom+ InsertRowCount 后开始的那行.
                        //3. copyStylesFromRow 不会把合并的单元格也弄过来(即,这个参数的功能不是格式刷)
                        if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyAll)
                        {
                            //if (insertRowFrom ==11)
                            //{
                            //    var aaa = 2;
                            //}
                            worksheet.InsertRow(insertRowFrom, insertRows); //用这个参数创建的excel,文件体积要小,插入速度没测试
                        }
                        else if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyStyleAndMergeCell)
                        {
                            if (dictConfig[nth].InsertRowStyle.NeedCopyStyles)
                            {
                                //在测试中,数据量 >= EPPlusConfig.MaxRow07/2-1  时,程序会抛异常, 这个数据量值仅做参考
                                //解决方案,分批插入, 且分批插入的 RowFrom 必须是第一次 InsertRow 的结尾行, 不然 第三遍循环:填充数据 会异常
                                //同时又发现了一个bug: worksheet.InsertRow 第三个参数 要满足 _rows + _copyStylesFromRow < EPPlusConfig.MaxRow07 , 但是_copyStylesFromRow 又是  _rowFrom + _rows 后开始数的行数 .nnd. 为了 防止报错, 我后面写了if-else 结果就是 后面新增的行没有样式

                                var insertRowsMax = (EPPlusConfig.MaxRow07 / 2 - 1) - 1;
                                if (insertRows >= insertRowsMax)
                                {
                                    var insertCount = insertRows / insertRowsMax;
                                    var mod = insertRows % insertRowsMax;
                                    int _rowFrom; int _rows; int _copyStylesFromRow;
                                    for (int i = 0; i < insertCount; i++)
                                    {
                                        _rowFrom = insertRowFrom + i * insertRowsMax;
                                        _rows = insertRowsMax;
                                        _copyStylesFromRow = _rowFrom + _rows;
                                        //防止报错, 结果就是 后面新增的行没有样式
                                        if (_rows + _copyStylesFromRow > EPPlusConfig.MaxRow07)
                                        {
                                            worksheet.InsertRow(_rowFrom, _rows);
                                        }
                                        else
                                        {
                                            worksheet.InsertRow(_rowFrom, _rows, _copyStylesFromRow);
                                        }
                                    }
                                    if (mod > 0)
                                    {
                                        _rowFrom = insertRowFrom + insertCount * insertRowsMax;
                                        _rows = mod;
                                        _copyStylesFromRow = lastSpaceLineRowNumber;
                                        //防止报错, 结果就是 后面新增的行没有样式
                                        if (_rows + _copyStylesFromRow > EPPlusConfig.MaxRow07)
                                        {
                                            worksheet.InsertRow(_rowFrom, _rows);
                                        }
                                        else
                                        {
                                            worksheet.InsertRow(_rowFrom, _rows, _copyStylesFromRow);
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

                        #region 第二遍循环:处理样式 (Height要自己单独处理)


                        if (dictConfig[nth].InsertRowStyle.Operation == InsertRowStyleOperation.CopyAll)
                        {
                            var configLine = $"{leftColStr}{lastSpaceLineRowNumber}:{rightColStr}{lastSpaceLineRowNumber}";

                            for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                            {
                                int destRow = dictDestRow[i];

                                //copy 好比格式刷, 这里只格式化配置行所在的表格部分.
                                //Copy 效率比 CopyStyleAndMergedCellFromConfigRow 慢差不多一倍(测试数据4w条,要4秒多, 用上面的是2秒多,且文件体积也要小很多 好像有50% ) 
                                worksheet.Cells[configLine].Copy(worksheet.Cells[$"{leftColStr}{destRow}:{rightColStr}{destRow}"]);

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

                        #endregion
                    }

                    //第三遍循环:填充数据
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
                            object val = row[colMapperName]; //33xxxx19941111xxxx
                            int destCol = configLineCellPoint[j].Col;
                            ExcelRange cells = worksheet.Cells[destRow, destCol];

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
                                    var isFillData_Title = fillMethod.SynchronizationDataSource.NeedTitle && i == 0;
                                    var isFillData_Body = fillMethod.SynchronizationDataSource.NeedBody;

                                    if (isFillData_Title || isFillData_Body)
                                    {
                                        if (fillDataColumnsStat == null)
                                        {
                                            fillDataColumnsStat = InitFillDataColumnStat(datatable, dictConfig[nth].ConfigLine, fillMethod);
                                        }

                                        if (isFillData_Title)
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
                                        if (isFillData_Body)
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
                            }

                            #endregion

                        }

                        if (config.IsReport)
                        {
                            SetReport(worksheet, row, config, destRow);
                        }
                    }
                }

                //已经修复bug:当只有一个配置时,且 deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白 原因是,配置行没被删除
                if (deleteLastSpaceLine)
                {
                    worksheet.DeleteRow(lastSpaceLineRowNumber, lastSpaceLineInterval, true);
                    sheetBodyAddRowCount -= lastSpaceLineInterval;
                }

                FillData_Body_Summary(config, worksheet, dictConfigSource, nth, dictConfig, sheetBodyAddRowCount);

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
                //if (item is ExcelCellRange && ((ExcelCellRange)item).IsMerge)
                //{
                //    rangeCells.Add(((ExcelCellRange)item));
                //}

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

            var allCell = new List<object>(); //所有的单元格
            //获得当前行,哪些单独单元格+合并单元格
            while (true)
            {
                if (worksheet.MergedCells[lineNo, leftAddressCol] == null)
                {
                    var cell = new ExcelCellPoint(new ExcelCellPoint(lineNo, leftAddressCol).R1C1);
                    allCell.Add(cell);
                    leftAddressCol++;
                }
                else
                {
                    var cell = new ExcelCellRange(new ExcelCellPoint(lineNo, leftAddressCol).R1C1, worksheet);
                    allCell.Add(cell);
                    leftAddressCol = cell.End.Col + 1;
                }

                if (leftAddressCol > rightAddressCol)
                {
                    break;
                }
            }

            return allCell;


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="worksheet"></param>
        /// <param name="dictConfigSource"></param>
        /// <param name="nth"></param>
        /// <param name="dictConfig"></param>
        /// <param name="sheetBodyAddRowCount"></param>
        private static void FillData_Body_Summary(EPPlusConfig config, ExcelWorksheet worksheet, Dictionary<int, EPPlusConfigSourceBodyOption> dictConfigSource, int nth, Dictionary<int, EPPlusConfigBodyOption> dictConfig, int sheetBodyAddRowCount)
        {

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
            var splitInclude = columns.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

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

        private static void FillData_Foot(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet, int sheetBodyAddRowCount)
        {
            //填充foot
            if (config.Foot.ConfigCellList == null || config.Foot.ConfigCellList.Count <= 0)
            {
                return;
            }

            var dictConfigSource = configSource.Foot.CellsInfoList.ToDictionary(a => a.ConfigValue);
            foreach (var item in config.Foot.ConfigCellList)
            {
                if (configSource.Foot == null || configSource.Foot.CellsInfoList == null ||
                    configSource.Foot.CellsInfoList.Count == 0) //excel中有配置foot,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
                {
                    break;
                }

                //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值
                var cellPoint = new ExcelCellPoint(item.Address);
                string colMapperName = item.ConfigValue;
                object val = dictConfigSource[item.ConfigValue].FillValue;
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
        /// 从Excel 中获得符合C# 类属性定义的列名集合
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">列名在Excel的第几行</param>
        /// <param name="colStart"></param>
        /// <param name="colEnd"></param>
        /// <param name="POCO_Property_AutoReame_WhenRepeat">当属性重复时自动重命名</param>
        /// <param name="renameFirtNameWhenRepeat">当属性重复时需要重命名第一个属性的名字</param>
        /// <returns></returns>
        private static List<ExcelCellInfo> GetExcelColumnOfModel(ExcelWorksheet ws, int row, int colStart, int? colEnd, bool POCO_Property_AutoReame_WhenRepeat = false, bool renameFirtNameWhenRepeat = true)
        {
            List<string> colNameList = null;
            Dictionary<string, int> nameRepeatCounter = null;
            if (POCO_Property_AutoReame_WhenRepeat)
            {
                colNameList = new List<string>();
                nameRepeatCounter = new Dictionary<string, int>();
            }
            if (colEnd == null) colEnd = EPPlusConfig.MaxCol07;
            var list = new List<ExcelCellInfo>();
            int col = colStart;
            while (col < colEnd)
            {
                int step;
                ExcelAddress ea;
                if (ws.Cells[row, col].Merge)
                {
                    ea = new ExcelAddress(ws.MergedCells[row, col]);
                    step = ea.Columns;
                }
                else
                {
                    ea = new ExcelAddress(row, col, row, col);
                    step = 1;
                }
                var colName = ws.Cells[ea.Start.Row, ea.Start.Column].Text;

                if (string.IsNullOrEmpty(colName)) break;
                colName = ExtractName(colName);
                if (string.IsNullOrEmpty(colName)) break;
                if (POCO_Property_AutoReame_WhenRepeat)
                {
                    AutoRename(colNameList, nameRepeatCounter, colName, renameFirtNameWhenRepeat);
                }
                list.Add(new ExcelCellInfo()
                {
                    WorkSheet = ws,
                    ExcelAddress = ea,
                    Value = colName,
                });

                col += step;

            }
            if (POCO_Property_AutoReame_WhenRepeat)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    var item = list[i];
                    item.Value = colNameList[i];
                }
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
            //去掉不合理的属性命名的字符串(提取合法的字符并接成一个字符串)
            string reg = @"[_a-zA-Z\u4e00-\u9FFF][A-Za-z0-9_\u4e00-\u9FFF]*";
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

            throw new System.Exception("GetList_SetModelValue()时遇到未处理的类型!!!请完善程序");
        }

        private static void GetList_ValidAttribute<T>(PropertyInfo pInfo, T model, string value) where T : class, new()
        {
            object[] validAttrs = ReflectionHelper.GetAttributeForProperty<T, ValidationAttribute>(pInfo.Name, true);
            if (validAttrs != null && validAttrs.Length > 0)
            {
                //同一个特性的的属性值肯定是一样的,所以可以优化;
                //ArrayList objArr = null;
                object[] objArr = null; //第二次优化
                foreach (var validAttr in validAttrs)
                {
                    MethodInfo methodIsValid = validAttr.GetType().GetMethod("IsValid");

                    #region 代码优化第二次,还是因为只有一个参数进行优化

                    ////var objArr = new ArrayList();
                    //var paras = methodIsValid.GetParameters();
                    ////ValidationAttribute的IsValid 只有一个Object的参数, 所以不需要判断 (但不绝对),如果自定义的存在多个,那么上面一行代码就会抛出异常:发现不明确的匹配。
                    ////if (paras.Length != 1)
                    ////{
                    ////    throw new Exception($@"遇到了在说");
                    ////}

                    //if (objArr == null)
                    //{
                    //    objArr = new ArrayList();
                    //    #region 只有一个参数,可以优化如下
                    //    objArr.Add(value);
                    //    //foreach (ParameterInfo paraInfo in paras)
                    //    //{
                    //    //    objArr.Add(value);
                    //    //    /*
                    //    //     *ValidationAttribute的IsValid 只有一个Object的参数,所以,直接Add就好了;
                    //    //   if (paraInfo.ParameterType.IsValueType)
                    //    //   {
                    //    //       //t.o.d.o...
                    //    //   }
                    //    //   else
                    //    //   {
                    //    //       objArr.Add(value);
                    //    //   }
                    //    //    */ 
                    //    //}  
                    //    #endregion
                    //} 
                    //var IsValid = (bool)methodIsValid.Invoke(validAttr, objArr.ToArray());

                    if (objArr == null)
                    {
                        objArr = new object[] { value };
                    }

                    var isValid = (bool)methodIsValid.Invoke(validAttr, objArr);

                    #endregion

                    if (!isValid)
                    {
                        var msg = $@"'{model.GetType().FullName}'类型的'{pInfo.Name}'属性验证未通过:'{((ValidationAttribute)validAttr).ErrorMessage}'";
                        throw new ArgumentException(msg);
                    }
                }
            }
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
                    throw new System.ArgumentException($"Value值:'{value}'在枚举值:'{pInfoType.FullName}'中未定义,请检查!!!");
                }

                throw new System.ArgumentException(attr.ErrorMessage);
            }

            //拼接ErrorMessage
            var allProp = ReflectionHelper.GetProperties<T>();

            for (int i = 0; i < attr.Args.Length; i++)
            {
                var propertyName = attr.Args[i];
                if (string.IsNullOrEmpty(propertyName))
                {
                    continue;
                }

                //如果占位符这是常量且刚好和属性名一直,请把占位符拆成多个占位符使用
                if (propertyName == pInfo.Name)
                {
                    attr.Args[i] = value;
                }
                else
                {
                    var prop = ReflectionHelper.GetProperty(allProp, propertyName, true);
                    if (prop == null)
                    {
                        continue;
                    }
                    attr.Args[i] = prop.GetValue(model).ToString();
                }
            }

            string message = string.Format(attr.ErrorMessage, attr.Args);
            throw new System.ArgumentException(message);

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

    }
}
