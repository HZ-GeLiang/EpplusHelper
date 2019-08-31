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
        public static List<string> FillDataWorkSheetNames = new List<string>();

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
            int sheetCount = excelPackage.Workbook.Worksheets.Count;
            if (workSheetIndex > sheetCount)
            {
                throw new ArgumentException($@"形参{nameof(workSheetIndex)}大于当前Excel的工作簿数量", nameof(workSheetIndex));//指定的参数已超出有效值的范围
            }
            return excelPackage.Workbook.Worksheets[workSheetIndex];
        }

        /// <summary>
        /// 根据workSheetIndex获得模版worksheet,然后复制一份出来并重命名成workSheetName并返回 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始</param>
        /// <param name="workSheetNewName"></param>
        /// <returns></returns>
        public static ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage, int workSheetIndex, string workSheetNewName)
        {
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
            var wsMom = GetExcelWorksheet(excelPackage, workSheetIndex);
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
        ///  
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
        /// <param name="workSheetNameExcludeList">要保留的工作簿名字</param>
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
            EPPlusHelper.FillDataWorkSheetNames.Add(workSheetNewName);
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
            ExcelWorksheet worksheet = EPPlusHelper.GetExcelWorksheet(excelPackage, destWorkSheetIndex, workSheetNewName);
            EPPlusHelper.FillDataWorkSheetNames.Add(workSheetNewName);
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

        private static int FillData_Body(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            //填充body
            var sheetBodyAddRowCount = 0; //新增了几行 (统计sheet body 在原有的模版上新增了多少行)

            if (config == null || configSource == null ||
                config.Body == null || configSource.Body == null ||
                config.Body.InfoList == null || configSource.Body.ConfigList == null ||
                config.Body.InfoList.Count <= 0 || configSource.Body.ConfigList.Count <= 0)
            {
                return sheetBodyAddRowCount;
            }

            int sheetBodyDeleteRowCount = 0; //sheet body 中删除了多少行(只含配置的行,对于FillData()内的删除行则不包括在内).  
            var dictConfig = config.Body.InfoList.ToDictionary(a => a.Nth, a => a.Option);
            var dictConfigSource = configSource.Body.ConfigList.ToDictionary(a => a.Nth, a => a.Option);
            foreach (var itemInfo in config.Body.InfoList)
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

                    if (!config.DeleteFillDateStartLineWhenDataSourceEmpty || dictConfig[nth].MapperExcel.Count <= 0)
                    {
                        continue; //跳过本次fillDate的循环
                    }

                    #region DeleteFillDateStartLine

                    foreach (var cellConfigInfo in dictConfig[nth].MapperExcel) //只遍历一次
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
                var hasMergeCell = dictConfig[nth].MapperExcel.Find(a => a.Address.Contains(":")) != null;
                Dictionary<string, FillDataColums> fillDataColumsStat = null;//Datatable 的列的使用情况

                if (hasMergeCell)
                {
                    //注:进入这里的条件是单元格必须是多行合并的,如果是同行多列合并的单元格,最后生成的excel会有问题,打开时会提示修复(修复完成后内容是正确的(不保证,因为我测试的几个内容是正确的))
                    List<ExcelCellRange> cellRange = dictConfig[nth].MapperExcel.Select(cellConfigInfo => new ExcelCellRange(cellConfigInfo.Address)).ToList();
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
                        for (int j = 0; j < dictConfig[nth].MapperExcel.Count; j++)
                        {
                            #region 赋值
                            string colMapperName = dictConfig[nth].MapperExcel[j].ConfigValue;
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
                                dictConfig[nth].CustomSetValue.Invoke(colMapperName, val, cells);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }
                            #endregion

                            #region 同步数据源
                            if (j == cellRange.Count - 1) //如果一行循环到了最后一列
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
                                    if ((isFillData_Title) || isFillData_Body)
                                    {
                                        if (fillDataColumsStat == null)
                                        {
                                            fillDataColumsStat = InitFillDataColumnStat(datatable, dictConfig[nth].MapperExcel, fillMethod);
                                        }

                                        if (isFillData_Title)
                                        {
                                            var eachCount = 0;
                                            var config_firstCell_col = new ExcelCellPoint(dictConfig[nth].MapperExcel.First().Address).Col;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {
                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                var extensionDestCol_title = config_firstCell_col + dictConfig[nth].MapperExcel.Count + eachCount;
                                                var extensionCell_Title = worksheet.Cells[destRow - 1, extensionDestCol_title];
                                                SetWorksheetCellsValue(config, extensionCell_Title, item.ColumName, item.ColumName);
                                                eachCount++;
                                            }
                                        }
                                        if (isFillData_Body)
                                        {
                                            var eachCount = 0;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {

                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                int extensionDestStartCol = cellRange[j].Start.Col + 1;
                                                int extensionDestEndCol = cellRange[j].End.Col + 1;
                                                var extensionCell = worksheet.Cells[destRow, extensionDestStartCol, destRow + maxIntervalRow, extensionDestEndCol];
                                                SetWorksheetCellsValue(config, extensionCell, row[item.ColumName], item.ColumName);
                                                eachCount++;
                                            }
                                        }
                                    }
                                }
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
                    var cellRange = dictConfig[nth].MapperExcel.Select(configCellsInfo => new ExcelCellPoint(configCellsInfo.Address)).ToList(); // 将配置的值 转换成 ExcelCellPoint

                    var cellFirst = cellRange.First();
                    var cellLast = cellRange.Last();

                    //这4个变量必须在 InsertRow之前运算
                    int cellFirstColStart = cellFirst.Col;
                    string cellFirstColStartZm = ExcelCellPoint.R1C1FormulasReverse(cellFirstColStart);
                    int cellLastColEnd = worksheet.Cells[cellLast.R1C1].Merge ? new ExcelCellRange(cellLast.R1C1, worksheet).End.Col : cellLast.Col;
                    string cellLastColEndZm = ExcelCellPoint.R1C1FormulasReverse(cellLastColEnd);

                    //第一遍循环:计算要插入多少行
                    var insertRowCount = 0;
                    var insertRowFrom = 0;
                    var dictDestRow = new Dictionary<int, int>();//数据源的第N行,对应excel填充的第N行
                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        int destRow = CalcDestRow(nth, sheetBodyAddRowCount, cellFirst, i, sheetBodyDeleteRowCount, currentLoopAddLines, cellRange);

                        dictDestRow.Add(i, destRow);
                        if (datatable.Rows.Count > 1) //1.数据源中的数据行数大于1才增行
                        {
                            if (i > tempLine - 2) //i从0开始,这边要-1,然后又要留一行模版,做为复制源,所以这里要-2  
                            {
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

                                insertRowCount++;

                                sheetBodyAddRowCount++;
                                currentLoopAddLines++;
                            }
                        }

                    }

                    if (insertRowCount > 0 && insertRowFrom > 0)
                    {
                        //在  InsertRowFrom 行前面插入 InsertRowCount 行.
                        //注:
                        //1. copyStylesFromRow 的行计算是在 InsertRowFrom+ InsertRowCount 后开始的那行.
                        //2. copyStylesFromRow 不会把合并的单元格也弄过来(即,这个参数的功能不是格式刷)
                        worksheet.InsertRow(insertRowFrom, insertRowCount, lastSpaceLineRowNumber);
                        //worksheet.InsertRow(insertRowFrom, insertRowCount); //用这个参数创建的excel,文件体积要小,插入速度没测试

                        #region 第二遍循环:处理样式 

                        var configLine_LineNo = lastSpaceLineRowNumber;
                        var configLine_r1c1_left = $"{cellFirstColStartZm}{configLine_LineNo}";
                        var configLine_r1c1_right = $"{cellLastColEndZm}{configLine_LineNo}";
                        var configLine = $"{configLine_r1c1_left}:{configLine_r1c1_right}";

                        #region 获得合并单元格

                        var Start = new ExcelCellPoint(configLine_r1c1_left);
                        var End = new ExcelCellPoint(configLine_r1c1_right);

                        var listCell = new List<object>();
                        var col = new ExcelCellPoint(configLine_r1c1_left).Col;
                        //获得当前行,哪些单独单元格+合并单元格
                        while (true)
                        {
                            if (worksheet.MergedCells[configLine_LineNo, col] == null)
                            {
                                var cell = new ExcelCellPoint(new ExcelCellPoint(configLine_LineNo, col).R1C1);
                                listCell.Add(cell);
                                col++;
                            }
                            else
                            {
                                var cell = new ExcelCellRange(new ExcelCellPoint(configLine_LineNo, col).R1C1, worksheet);
                                listCell.Add(cell);
                                col = cell.End.Col + 1;
                            }
                            if (col > End.Col)
                            {
                                break;
                            }
                        }

                        var rangeCells = new List<ExcelCellRange>();
                        foreach (var item in listCell)
                        {
                            if (item is ExcelCellRange && ((ExcelCellRange)item).IsMerge)
                            {
                                rangeCells.Add(((ExcelCellRange)item));
                            }
                        }

                        #endregion

                        for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                        {
                            int destRow = dictDestRow[i];

                            #region 不用copy,只 合并单元格 关于背景颜这些,InsertRow 使用第三个参数就可以了
                            foreach (var item in rangeCells)
                            {
                                //var r1c1 = $"{item.Start.ColStr}{destRow}:{item.End.ColStr}{destRow}";
                                ////if (!worksheet.Cells[r1c1].Merge)
                                ////{
                                ////     worksheet.Cells[r1c1].Merge = true; 
                                ////}
                                //worksheet.Cells[r1c1].Merge = true; //insert row 的数据没有合并单元格的, 所有就去掉了判断

                                //不用r1c1, 不优化程序, 节省一步创建字符串的过程
                                worksheet.Cells[destRow, item.Start.Col, destRow, item.End.Col].Merge = true;
                            }
                            #endregion

                            #region 使用copy
                            //var destRowLine = $"{cellFirstColStartZm}{destRow}:{cellLastColEndZm}{destRow}";
                            //worksheet.Cells[configLine].Copy(worksheet.Cells[destRowLine]);//copy好比格式刷, 这里只格式化配置行所在的表格部分. 效率比上面注释的慢差不多一倍(测试数据4w条,要4秒多, 用上面的是2秒多,且文件体积也要小 50% ) 
                            #endregion

                            //不要用[row,col]索引器,[row,col]表示某单元格.注意:copy会把source行的除了height(觉得是一个bug)以外的全部复制一行出来
                            worksheet.Row(destRow).Height = worksheet.Row(configLine_LineNo).Height; //修正height


                        }

                        #endregion
                    }

                    //第三遍喜欢:填充数据
                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        int destRow = dictDestRow[i];
                        DataRow row = datatable.Rows[i];

                        //3.赋值.
                        //注:遍历时变量 j 的终止条件不能是 datatable.Rows.Count. 因为datatable可能会包含多余的字段信息,与 配置信息列的个数不一致.
                        for (int j = 0; j < dictConfig[nth].MapperExcel.Count; j++)
                        {
                            #region 赋值

                            //worksheet.Cells[destRow, destCol].Value = row[j];
                            string colMapperName = dictConfig[nth].MapperExcel[j].ConfigValue;
                            object val = row[colMapperName];
                            int destCol = cellRange[j].Col;
                            ExcelRange cells = worksheet.Cells[destRow, destCol];

                            if (dictConfig[nth].CustomSetValue != null)
                            {
                                dictConfig[nth].CustomSetValue.Invoke(colMapperName, val, cells);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }

                            #endregion

                            #region 同步数据源

                            if (j == cellRange.Count - 1) //如果一行循环到了最后一列
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
                                        if (fillDataColumsStat == null)
                                        {
                                            fillDataColumsStat = InitFillDataColumnStat(datatable, dictConfig[nth].MapperExcel, fillMethod);
                                        }

                                        if (isFillData_Title)
                                        {
                                            var eachCount = 0;
                                            var config_firstCell_col = new ExcelCellPoint(dictConfig[nth].MapperExcel.First().Address).Col;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {
                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                var extensionDestCol_title = config_firstCell_col + dictConfig[nth].MapperExcel.Count + eachCount;
                                                var extensionCell_Title = worksheet.Cells[destRow - 1, extensionDestCol_title];
                                                SetWorksheetCellsValue(config, extensionCell_Title, item.ColumName, item.ColumName);
                                                eachCount++;
                                            }
                                        }
                                        if (isFillData_Body)
                                        {
                                            var eachCount = 0;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {
                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                var extensionDestCol = cellRange[j].Col + 1 + eachCount;

                                                var extensionCell = worksheet.Cells[destRow, extensionDestCol];
                                                SetWorksheetCellsValue(config, extensionCell, row[item.ColumName], item.ColumName);
                                                eachCount++;
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
            }

            return sheetBodyAddRowCount;
        }

        /// <summary>
        /// 填充第N个配置的一些零散的单元格的值(譬如汇总信息等)
        /// </summary>
        /// <param name="config"></param>
        /// <param name="worksheet"></param>
        /// <param name="dictConfigSource"></param>
        /// <param name="nth"></param>
        /// <param name="dictConfig"></param>
        /// <param name="sheetBodyAddRowCount"></param>
        private static void FillData_Body_Summary(EPPlusConfig config, ExcelWorksheet worksheet, Dictionary<int, EPPlusConfigSourceBodyOption> dictConfigSource, int nth, Dictionary<int, EPPlusConfigBodyOption> dictConfig, int sheetBodyAddRowCount)
        {
            if (dictConfigSource[nth].Summary == null) return;
            var dictConfigSourceSummary = dictConfigSource[nth].Summary.ToDictionary(a => a.ConfigValue);
            foreach (var item in dictConfig[nth].SummaryMapperExcel)
            {
                var excelCellPoint = new ExcelCellPoint(item.Address);
                string colMapperName = item.ConfigValue;
                object val = dictConfigSourceSummary[colMapperName].FillValue;
                ExcelRange cells = worksheet.Cells[excelCellPoint.Row + sheetBodyAddRowCount, excelCellPoint.Col];

                if (dictConfig[nth].SummaryCustomSetValue != null)
                {
                    dictConfig[nth].SummaryCustomSetValue.Invoke(colMapperName, val, cells);
                }
                else
                {
                    SetWorksheetCellsValue(config, cells, val, colMapperName);
                }
            }
        }

        private static int CalcDestRow(int nth, int sheetBodyAddRowCount, ExcelCellPoint fillData_FirstCellInfo, int i,
            int sheetBodyDeleteRowCount, int currentLoopAddLines, List<ExcelCellPoint> startCellPointLine)
        {
            int destRow;
            if (nth == 1)
            {
                //destRow = sheetBodyAddRowCount > 0
                //? startCellPointLine[0].Row + i - sheetBodyDeleteRowCount
                //: startCellPointLine[0].Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                destRow = sheetBodyAddRowCount > 0
                    ? fillData_FirstCellInfo.Row + i - sheetBodyDeleteRowCount
                    : fillData_FirstCellInfo.Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
            }
            else
            {
                //destRow = currentLoopAddLines > 0
                //    ? startCellPointLine[0].Row + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                //    : startCellPointLine[0].Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;

                destRow = currentLoopAddLines > 0
                    ? fillData_FirstCellInfo.Row + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                    : fillData_FirstCellInfo.Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
            }

            return destRow;
        }

        /// <summary>
        ///  获得Database数据源的所有的列的使用状态
        /// </summary>
        /// <param name="datatable"></param>
        /// <param name="nth"></param>
        /// <param name="fillModel"></param>
        /// <returns></returns>
        private static Dictionary<string, FillDataColums> InitFillDataColumnStat(DataTable datatable, List<EPPlusConfigFixedCell> nth, SheetBodyFillDataMethod fillModel)
        {
            var fillDataColumnStat = new Dictionary<string, FillDataColums>();
            foreach (DataColumn column in datatable.Columns)
            {
                fillDataColumnStat.Add(column.ColumnName, new FillDataColums()
                {
                    ColumName = column.ColumnName,
                    State = FillDataColumsState.Unchanged
                });
            }

            foreach (var item in nth)
            {
                fillDataColumnStat[item.ConfigValue].State = FillDataColumsState.Used;
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

        private static void Modify_DataColumnsIsUsedStat(Dictionary<string, FillDataColums> fillDataColumsStat, string columns, bool selectColumnIsWillUse)
        {
            var splitInclude = columns.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var key in fillDataColumsStat.Keys)
            {
                if (fillDataColumsStat[key].State == FillDataColumsState.Unchanged)
                {
                    if (splitInclude.Contains(key))
                    {
                        fillDataColumsStat[key].State = selectColumnIsWillUse ? FillDataColumsState.WillUse : FillDataColumsState.WillNotUse;
                    }
                    else
                    {
                        fillDataColumsStat[key].State = selectColumnIsWillUse ? FillDataColumsState.WillNotUse : FillDataColumsState.WillUse;
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
                //var cellpoint = R1C1ToExcelCellPoint(item.Key);
                var cellpoint = new ExcelCellPoint(item.Address);
                // worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col].Value = configSource.SheetFoot[item.Value];
                string colMapperName = item.ConfigValue;
                object val = dictConfigSource[item.ConfigValue].FillValue;
                ExcelRange cells = worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col];
                if (config.Foot.CellCustomSetValue != null)
                {
                    config.Foot.CellCustomSetValue.Invoke(colMapperName, val, cells);
                }
                else
                {
                    //item.Key -> D13 , item.Value -> 总计
                    SetWorksheetCellsValue(config, cells, val, colMapperName);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="cells">s结尾表示单元格有可能是合并单元格</param>
        /// <param name="val">值</param>
        /// <param name="colMapperName">excel填充的列名,不想传值请使用null</param> 
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
            int step;
            while (col < colEnd)
            {
                ExcelAddress ea;
                string colName;
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
                colName = ws.Cells[ea.Start.Row, ea.Start.Column].Text;

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
            int level = Convert.ToInt32(row[config.Report.RowLevelColumnName]) - 1;//leve是从0开始的
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
                    //    //       //todo:...
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
                InfoList = new List<EPPlusConfigBodyInfo>()
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

        #region GetList<T>

        public static GetExcelListArgs<T> GetExcelListArgsDefault<T>(ExcelWorksheet ws, int rowIndex) where T : class
        {
            return new GetExcelListArgs<T>()
            {
                ws = ws,
                RowIndex_Data = rowIndex,
                EveryCellPrefix = "",
                EveryCellReplaceList = null,
                RowIndex_DataName = rowIndex - 1,
                UseEveryCellReplace = true,
                HavingFilter = null,
                WhereFilter = null,
                ReadCellValueOption = ReadCellValueOption.Trim,
                POCO_Property_AutoRename_WhenRepeat = false,
                POCO_Property_AutoRenameFirtName_WhenRepeat = true,
                ScanLine = ScanLine.MergeLine,
                MatchingModelEqualsCheck = true,
                GetList_NeedAllException = false,
                GetList_ErrorMessage_OnlyShowColomn = false,
            };
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
        /// <param name="everyCellPrefix">被遍历的单元格内容不为空时的起始字符必须是该字符,然后忽略该字符</param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellPrefix, string everyCellReplaceOldValue, string everyCellReplaceNewValue) where T : class, new()
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

        public static List<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.RowIndex_Data;
            if (rowIndex <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowIndex}'错误,值应该大于0");
            }

            int rowIndex_DataName = args.RowIndex_DataName;
            if (rowIndex_DataName <= 0)
            {
                throw new ArgumentException($@"数据起始行的标题行值'{rowIndex_DataName}'错误,值应该大于0");
            }

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, 1, EPPlusConfig.MaxCol07, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
            if (colNameList.Count == 0)
            {
                throw new Exception("未读取到单元格标题");
            }
            var dictExcelAddressCol = colNameList.ToDictionary(item => item.ExcelAddress, item => new ExcelCellPoint(item.ExcelAddress).Col);

            Type type = typeof(T);

            #region 获得字典
            var dictModelPropNameExistsExcelColumn = new Dictionary<string, bool>();//Model属性在Excel列中存在, key: ModelPropName
            var dictModelPropNameToExcelColumnName = new Dictionary<string, string>();//Model属性名字对应的excel的标题列名字
            var dictExcelColumnIndexToModelPropName_Temp = new Dictionary<int, string>();//Excel的列标题和Model属性名字的映射
            var dictExcelColumnIndexToModelPropName_All = new Dictionary<int, string>();//Excel列对应的Model属性名字(所有excel列)

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
                        PropertyInfo pInfoTemp = null;
                        var propNameTemp = dictExcelColumnIndexToModelPropName_Temp[excelColumnIndex];
                        //不做属性的 DisplayExcelColumnName = 当前属性的验证 (因为还没想到这个属性是一定要验证的情况)
                        //if (dictModelPropNameToExcelColumnName.ContainsKey(propNameTemp) && dictModelPropNameToExcelColumnName[propNameTemp] == propName)
                        //{
                        //   pInfoTemp = type.GetProperty(propName);
                        //}

                        pInfoTemp = type.GetProperty(propNameTemp);
                        if (pInfoTemp != null)
                        {
                            propName = propNameTemp;
                            pInfo = pInfoTemp;
                        }
                    }
                }

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
                        dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, ws));
                    }
                    else if (matchingModelValue == MatchingModel.lt)
                    {
                        if ((_matchingModel & MatchingModel.lt) != MatchingModel.lt) continue;
                        if (dictMatchingModelException.ContainsKey(MatchingModel.lt)) continue;
                        dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, ws));
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
                            dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, ws));
                        }
                        if (!dictMatchingModelException.ContainsKey(MatchingModel.lt))
                        {
                            dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, ws));
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
                ? GetExcelListArgs<T>.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;

            var ctor = type.GetConstructor(new Type[] { });
            if (ctor == null) throw new ArgumentException($"通过反射无法得到'{type.FullName}'的一个无构造参数的构造器.");

            var dictPropAttrs = new Dictionary<string, Dictionary<string, Attribute>>();//属性里包含的Attribute

            //内置的Attribute
            var dictUnique = new Dictionary<string, Dictionary<string, bool>>();//属性的 UniqueAttribute
            string key_UniqueAttribute = typeof(UniqueAttribute).FullName;

            foreach (ExcelCellInfo excelCellInfo in colNameList)
            {
                int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                {
                    continue;
                }
                string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                PropertyInfo pInfo = type.GetProperty(propName);
                if (pInfo == null)//防御式编程判断
                {
                    throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                }

                #region 初始化Attr要处理相关的数据
                var attrDict = new Dictionary<string, Attribute>() { };//key 是Attribute的FullName
                dictPropAttrs.Add(pInfo.Name, attrDict);

                var uniqueAttrs = ReflectionHelper.GetAttributeForProperty<UniqueAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (uniqueAttrs.Length > 0)
                {
                    dictPropAttrs[pInfo.Name].Add(key_UniqueAttribute, (UniqueAttribute)uniqueAttrs[0]);
                    dictUnique.Add(pInfo.Name, new Dictionary<string, bool>());
                }

                #endregion
            }

            #region 获得 list
            List<T> list = new List<T>();
            int row = rowIndex;
            Exception exception = null;

            int? step = null;
            switch (args.ScanLine)
            {
                case ScanLine.SingleLine:
                    step = 1;
                    break;
                case ScanLine.MergeLine:
                    //while里面动态计算
                    break;
                default:
                    throw new Exception("不支持的ScanLine");
            }

            var excelCellInfoNeedTrim = (args.ReadCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            var excelCellInfoNeedMergeLine = (args.ReadCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            var excelCellInfoNeedToDBC = (args.ReadCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;

            var allException = args.GetList_NeedAllException ? new List<Exception>() : null;
            while (true)
            {
                bool isNoDataAllColumn = true;//判断整行数据是否都没有数据
                T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                    if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                    {
                        continue;
                    }
                    string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                    if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                    PropertyInfo pInfo = type.GetProperty(propName);
                    if (pInfo == null)//防御式编程判断
                    {
                        throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                    }

                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

#if DEBUG
                    string value;
                    if (pInfo.PropertyType == typeof(DateTime?) || pInfo.PropertyType == typeof(DateTime))
                    {
                        //todo:对于日期类型的,有时候要获取Cell.Value, 有空了修改
                        value = GetMergeCellText(ws, row, col);
                    }
                    else
                    {
                        value = GetMergeCellText(ws, row, col);
                    }
#else
                    string value =  GetMegerCellText(ws, row, col);
#endif

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
                                throw new System.ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
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

                        if (excelCellInfoNeedTrim)
                        {
                            value = value.Trim();
                        }
                        if (excelCellInfoNeedMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (excelCellInfoNeedToDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion

                        #region 处理内置的Attribute

                        if (dictPropAttrs[pInfo.Name].ContainsKey(key_UniqueAttribute))
                        {
                            var uniqueAttr = (UniqueAttribute)dictPropAttrs[pInfo.Name][key_UniqueAttribute];
                            var uniqueAttrAttr_ErrorMsg_IsNullOrEmpty = string.IsNullOrEmpty(uniqueAttr.ErrorMessage);
                            if (!valueIsNullOrEmpty)
                            {
                                if (!dictUnique[pInfo.Name].ContainsKey(value))
                                {
                                    dictUnique[pInfo.Name].Add(value, default(bool));
                                }
                                else
                                {
                                    string exception_msg = uniqueAttrAttr_ErrorMsg_IsNullOrEmpty ? $@"属性'{pInfo.Name}'的值:'{value}'出现了重复" : uniqueAttr.ErrorMessage;
                                    throw new ArgumentException(exception_msg, pInfo.Name);
                                }
                            }

                        }
                        #endregion
                    }
                    try
                    {
                        //验证特性
                        GetList_ValidAttribute(pInfo, model, value);
                        //赋值
                        GetList_SetModelValue(pInfo, model, value);
                    }
                    catch (Exception e)
                    {
                        exception = e is ArgumentException ? new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}", e) : e;
                        break;
                    }
                }

                if (isNoDataAllColumn)
                {
                    if (row == rowIndex)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    break; //出现空行,读取模版结束
                }

                //先添加Step
                if (step != null)
                {
                    row += (int)step;
                }
                else
                {
                    string range = ws.MergedCells[row, 1];
                    if (range == null)
                    {
                        row += 1;
                    }
                    else
                    {
                        var ea = new ExcelAddress(range);
                        row += ea.Rows;
                    }
                }
                //在判断异常

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

                sb.RemoveLastChar(',');
                throw new ArgumentException(sb.ToString());
            }

            #endregion

            return args.HavingFilter == null ? list : list.Where(item => args.HavingFilter.Invoke(item)).ToList();
        }

        public static DataTable GetDataTable(GetExcelListArgs<DataRow> args)
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.RowIndex_Data;
            if (rowIndex <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowIndex}'错误,值应该大于0");
            }

            int rowIndex_DataName = args.RowIndex_DataName;
            if (rowIndex_DataName <= 0)
            {
                throw new ArgumentException($@"数据起始行的标题行值'{rowIndex_DataName}'错误,值应该大于0");
            }

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, 1, EPPlusConfig.MaxCol07, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
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
                ? GetExcelListArgs<DataRow>.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;


            #region 获得 list

            int row = rowIndex;
            int? step = null;
            switch (args.ScanLine)
            {
                case ScanLine.SingleLine:
                    step = 1;
                    break;
                case ScanLine.MergeLine:
                    //while里面动态计算
                    break;
                default:
                    throw new Exception("不支持的ScanLine");
            }

            var excelCellInfoNeedTrim = (args.ReadCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            var excelCellInfoNeedMergeLine = (args.ReadCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            var excelCellInfoNeedToDBC = (args.ReadCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;

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
                                throw new System.ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
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

                        if (excelCellInfoNeedTrim)
                        {
                            value = value.Trim();
                        }
                        if (excelCellInfoNeedMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (excelCellInfoNeedToDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion

                    }

                    //赋值
                    dr[propName] = value;
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

                if (step != null)
                {
                    row += (int)step;
                }
                else
                {
                    string range = ws.MergedCells[row, 1];
                    if (range == null)
                    {
                        row += 1;
                    }
                    else
                    {
                        var ea = new ExcelAddress(range);
                        row += ea.Rows;
                    }
                }
            }

            #endregion

            return args.HavingFilter == null ? dt : dt.AsEnumerable().Where(item => args.HavingFilter.Invoke(item)).CopyToDataTable();
        }

        private static string DealMatchingModelException(MatchingModelException matchingModelException)
        {
            //注:这里的仅针对 MatchingModel.eq
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
            else if ((matchingModelException.MatchingModel & MatchingModel.gt) == MatchingModel.gt)
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
            else if ((matchingModelException.MatchingModel & MatchingModel.lt) == MatchingModel.lt)
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
            else if ((matchingModelException.MatchingModel & MatchingModel.neq) == MatchingModel.neq)
            {
                StringBuilder sb = new StringBuilder();
                return sb.ToString();
            }
            else
            {
                throw new Exception($@"参数{nameof(matchingModelException)},不支持的MatchingMode值");
            }

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
        private static MatchingModelException GetMatchingModelExceptionCase_gt(List<string> modelPropNotExistsExcelColumn,
            Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
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
            Dictionary<int, string> dictExcelColumnIndexToModelPropName_All,
            Dictionary<string, bool> dictModelPropNameExistsExcelColumn,
            out List<string> modelPropNotExistsExcelColumn, out List<string> excelColumnIsNotModelProp)
        {
            if (dictExcelColumnIndexToModelPropName_All == null) throw new ArgumentNullException(nameof(dictExcelColumnIndexToModelPropName_All));
            if (dictModelPropNameExistsExcelColumn == null) throw new ArgumentNullException(nameof(dictModelPropNameExistsExcelColumn));

            modelPropNotExistsExcelColumn = new List<string>();//model属性不在excel列中
            excelColumnIsNotModelProp = new List<string>();//excel列不是model属性

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.eq;
            }

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count > 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.neq | MatchingModel.gt;
            }

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count > 0)
            {
                return MatchingModel.neq | MatchingModel.lt;
            }


            foreach (var excelColumnIndex in dictExcelColumnIndexToModelPropName_All.Keys)
            {
                if (dictExcelColumnIndexToModelPropName_All[excelColumnIndex] == null)
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

        #region 一些帮助方法

        /// <summary>
        /// 获得合并单元格的值 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellText(ExcelWorksheet ws, int row, int col)
        {
            string range = ws.MergedCells[row, col];
            if (range == null) return GetCellText(ws, row, col);
            var ea = new ExcelAddress(range);
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
                return cell.Text;//有的单元格通过cell.Text取值会发生异常,但cell.Value却是有值的
            }
            catch (System.NullReferenceException e)
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
            //var dict = new Dictionary<string, string>() { { "A1", "序号" } };
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
        /// <param name="excelPackage"></param>
        /// <param name="sheetTitleLineNumber">工作簿标题行,key:工作簿名字,value:行号</param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, Dictionary<string, int> sheetTitleLineNumber)
        {
            if (sheetTitleLineNumber == null)
            {
                sheetTitleLineNumber = new Dictionary<string, int>();
            }
            ExcelWorksheets wss = excelPackage.Workbook.Worksheets;
            List<DefaultConfig> list = new List<DefaultConfig>();
            foreach (var ws in wss)
            {
                int titleLine;
                if (sheetTitleLineNumber != null || sheetTitleLineNumber.ContainsKey(ws.Name))
                {
                    titleLine = sheetTitleLineNumber[ws.Name];
                }
                else
                {
                    titleLine = 2;
                }

                list.Add(FillExcelDefaultConfig(ws, titleLine));
            }

            return list;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileOutDirectoryName"></param>
        /// <param name="sheetTitleLineNumber"></param>
        /// <param name="cellCustom"></param>
        /// <returns></returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(string filePath, string fileOutDirectoryName, Dictionary<int, int> sheetTitleLineNumber = null, Action<ExcelRange> cellCustom = null)
        {
            List<DefaultConfig> defaultConfigList;
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                defaultConfigList = FillExcelDefaultConfig(excelPackage, sheetTitleLineNumber, cellCustom);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save($@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result.xlsx");
            }
            return defaultConfigList;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="sheetTitleLineNumber">工作簿标题行,key:第几个工作簿,从1开始,value:行号</param>
        /// <param name="cellCustom"></param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, Dictionary<int, int> sheetTitleLineNumber, Action<ExcelRange> cellCustom = null)
        {
            ExcelWorksheets wss = excelPackage.Workbook.Worksheets;
            List<DefaultConfig> list = new List<DefaultConfig>();
            var eachCount = 0;
            foreach (var ws in wss)
            {
                int titleLine = sheetTitleLineNumber == null
                    ? 1
                    : sheetTitleLineNumber.ContainsKey(eachCount) ? sheetTitleLineNumber[eachCount] : 1;
                list.Add(FillExcelDefaultConfig(ws, titleLine, cellCustom));
                eachCount++;
            }
            return list;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="titleLineNumber"></param>
        /// <param name="cellCustom">对单元格进行额外处理</param>
        /// <returns></returns>
        public static DefaultConfig FillExcelDefaultConfig(ExcelWorksheet ws, int titleLineNumber, Action<ExcelRange> cellCustom = null)
        {
            var colNameList = new List<ExcelCellInfoValue>();
            var nameRepeatCounter = new Dictionary<string, int>();
            #region 获得colNameList

            int col = 1;
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

                if (ws.Cells[titleLineNumber, col].Merge)
                {
                    var address = ws.MergedCells[titleLineNumber, col];
                    var range = new ExcelCellRange(address);
                    col += range.IntervalCol + 1;
                }
                else
                {
                    col++;
                }
            }

            #endregion

            #region 给单元格赋值

            int fillBodyLine;
            if (ws.Cells[titleLineNumber, 1].Merge)
            {
                var address = ws.MergedCells[titleLineNumber, 1];
                var range = new ExcelCellRange(address);
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

            #region 关键字
            var columnTypeList_DateTime = new List<string>()
            {
                "时间", "日期", "date", "time","今天","昨天","明天","前天","day"
            };
            var columnTypeList_String = new List<string>()
            {
                "id","身份证","银行卡","卡号","手机","mobile","tel","序号","number","编号","No"
            };
            #endregion

            #region 关键字tolower
            for (int i = 0; i < columnTypeList_DateTime.Count; i++)
            {
                columnTypeList_DateTime[i] = columnTypeList_DateTime[i].ToLower();
            }
            for (int i = 0; i < columnTypeList_String.Count; i++)
            {
                columnTypeList_String[i] = columnTypeList_String[i].ToLower();
            }
            #endregion

            foreach (var colName in colNameList)
            {
                var propName = colName.IsRename ? colName.NameNew : colName.Name;
                sbColumn.AppendLine($"dt.Columns.Add(\"{propName}\");");
                sbAddDr.AppendLine($"//dr[\"{propName}\"] = ");

                var propName_lower = propName.ToLower();
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
                foreach (var item in columnTypeList_DateTime)
                {
                    if (propName_lower.IndexOf(item, StringComparison.Ordinal) != -1)
                    {
                        sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(DateTime);");
                        sb_CreateClassSnippet.AppendLine($" public DateTime {propName} {{ get; set; }}");
                        sb_CrateClassSnippe_AppendLine_InForeach = true;
                        break;
                    }
                }

                foreach (var item in columnTypeList_String)
                {
                    if (propName_lower.IndexOf(item, StringComparison.Ordinal) != -1)
                    {
                        sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(String);");
                        sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                        sb_CrateClassSnippe_AppendLine_InForeach = true;
                        break;//处理过了就break,不然会重复处理 譬如 银行卡号, 此时符合 银行卡 和卡号
                    }
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
            var cellsValue = ws.Cells.Value as object[,];
            if (cellsValue == null) throw new ArgumentNullException();

            return GetCellsBy(ws, cellsValue, a => a != null && a.ToString() == value);
        }

        /// <summary>
        /// 根据值获的excel中对应的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="cellsValue">一般通过ws.Cells.Value as object[,] 获得 </param>
        /// <param name="match">示例: a => a != null && a.GetType() == typeof(string) && ((string) a == "备注")</param>
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
                for (int workSheetIndex = 1; workSheetIndex <= excelPackage.Workbook.Worksheets.Count; workSheetIndex++)
                {
                    var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
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
                for (int workSheetIndex = 1; workSheetIndex <= excelPackage.Workbook.Worksheets.Count; workSheetIndex++)
                {
                    var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
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
            using (var fs = File.OpenRead(fileFullPath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                return ScientificNotationFormatToString(excelPackage, fileSaveAsPath, containNoMatchCell);
            }
        }

        #endregion

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
            var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
            EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, sheet);
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
            var sheet = GetExcelWorksheet(excelPackage, workSheetName);
            EPPlusHelper.SetDefaultConfigFromExcel(excelPackage, config, sheet);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, ExcelWorksheet sheet)
        {
            //让 sheet.Cells.Value 强制从A1单元格开始
            //遇到问题描述:创建一个excel,在C7,C8,C9,10单元格写入一些字符串, sheet.Cells.Value 是object[4,3]的数组, 但我要的是object[10,3]的数组
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value == null)
            {
                cellA1.Value = null;
            }

            config.Head = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$th") };
            SetConfigBodyFromExcel(config, sheet);
            config.Foot = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$tf") };
        }

        /// <summary>
        /// 设置sheetBody配置
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetConfigBodyFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");
            var sheetMergedCellsList = sheet.MergedCells.ToList();

            var dictList = new List<List<EPPlusConfigFixedCell>>();
            var dictSummeryList = new List<List<EPPlusConfigFixedCell>>();
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (cellStr.Length < 3) //配置至少有4个字符.所以,4个以下的都可以跳过
                    {
                        continue; //不用""比较,.Length速度比较快
                    }

                    if (!cellStr.StartsWith("$tb")) continue;

                    //  {"L15", "付款对象"}, $tb1
                    string cellPosition = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);

                    string nthStr = RegexHelper.GetFirstNumber(cellStr);
                    int nth = Convert.ToInt32(nthStr);
                    if (cellStr.StartsWith("$tbs")) //模版摘要/汇总等信息单元格
                    {
                        string cellConfigValue = Regex.Replace(cellStr, "^[$]tbs" + nthStr, ""); //$需要转义
                        if (dictSummeryList.Count < nth)
                        {
                            dictSummeryList.Add(new List<EPPlusConfigFixedCell>());
                        }

                        if (dictSummeryList[nth - 1].Find(a => a.ConfigValue == cellConfigValue) !=
                            default(EPPlusConfigFixedCell))
                        {
                            throw new ArgumentException($"Excel文件中的$tbs{nth}部分配置了相同的项:{cellConfigValue}");
                        }

                        dictSummeryList[nth - 1].Add(new EPPlusConfigFixedCell()
                        { Address = cellPosition, ConfigValue = cellConfigValue.Trim() });
                    }
                    else if (cellStr.StartsWith($"$tb{nthStr}$")) //模版提供了多少行,若没有配置,在调用FillData()时默认提供1行  $tb1$1
                    {
                        string cellConfigValue = Regex.Replace(cellStr, $@"^[$]tb{nth}[$]", ""); //$需要转义, 这个值一般都是数字

                        if (!int.TryParse(cellConfigValue, out int cellConfigValueInt))
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

                        var nthOption = config.Body.InfoList.Find(a => a.Nth == nth).Option;
                        if (nthOption.MapperExcelTemplateLine != null)
                        {
                            throw new ArgumentException($"Excel文件中重复配置了项$tb{nthStr}${cellConfigValue}");
                        }

                        nthOption.MapperExcelTemplateLine = cellConfigValueInt;
                    }
                    else //StartsWith($"$tb{nthStr}")
                    {
                        string cellConfigValue = Regex.Replace(cellStr, "^[$]tb" + nthStr, ""); //$需要转义

                        if (dictList.Count < nth)
                        {
                            dictList.Add(new List<EPPlusConfigFixedCell>());
                        }

                        if (dictList[nth - 1].Find(a => a.ConfigValue == cellConfigValue) !=
                            default(EPPlusConfigFixedCell))
                        {
                            throw new ArgumentException($"Excel文件中的$tb{nth}部分配置了相同的项:{cellConfigValue}");
                        }

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
                                //描述出现null的情况
                                /*
                                 * F10 G10
                                 * F11 G11
                                 * F12 G12
                                 * 这些单元格被合并为一个单元格,即用F10:G12来描述
                                 * 此时,配置单元格读取应该是F10,G10将不会被读取,
                                 * 直到上面为止,都是正确的,但是,偏偏有一个神一样的操作,
                                 * 当excel模版出现不规范操作(Excel一眼看上去将没有问题),G10单元格被读取出来后,那么在sheetMergedCellsList中肯定找不到
                                 * 然后下面一行代码就抛出未将对象引用设置到对象的实例异常
                                 * 该操作是:B10, D10, F10, G10单元格均有配置项,B10:C12进行单元格合并,然后用格式刷,对D10:E12, F10:G10进行格式化
                                 * 结果就是G10的配置项将会被隐藏,excel会提示 合并单元格时，仅保留左上角的值，而放弃其他值。但是这个其他值没有被清空,而是看不到了
                                 * 如果手动的合并F10:G10,Excel将会alert此操作会仅保留左上角的值
                                 */
                                throw new Exception($"工作簿{sheet.Name}的单元格{cellPosition}存在配置问题,请检查({cellPosition}是合并单元格,请取消合并,并把单元格的值给清空,然后重新合并)");
                            }

                            var cells = newKey.Split(':');

                            // ReSharper disable once ConvertIfStatementToConditionalTernaryExpression
                            if (RegexHelper.GetFirstNumber(cells[0]) == RegexHelper.GetFirstNumber(cells[1])) //是同一行的
                            {
                                dictList[nth - 1].Add(new EPPlusConfigFixedCell()
                                { Address = cellPosition, ConfigValue = cellConfigValue });
                            }
                            else
                            {
                                dictList[nth - 1].Add(new EPPlusConfigFixedCell()
                                { Address = newKey, ConfigValue = cellConfigValue });
                            }
                        }
                        else
                        {
                            dictList[nth - 1].Add(new EPPlusConfigFixedCell()
                            { Address = cellPosition, ConfigValue = cellConfigValue });
                        }
                    }

                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉. 注用这种方式清空值,,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                    sheet.Cells[i + 1, j + 1].Value = null; //修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白(后期已经修复)
                }
            }

            for (int i = 0; i < dictList.Count; i++)
            {
                var bodyInfo = config.Body.InfoList.Find(a => a.Nth == i + 1);
                if (bodyInfo == null)
                {
                    bodyInfo = new EPPlusConfigBodyInfo()
                    {
                        Nth = i + 1,
                        Option = new EPPlusConfigBodyOption(),
                    };
                    config.Body.InfoList.Add(bodyInfo);
                }

                bodyInfo.Option.MapperExcel = dictList[i];
            }

            for (int i = 0; i < dictSummeryList.Count; i++)
            {
                var bodyInfo = config.Body.InfoList.Find(a => a.Nth == i + 1);
                if (bodyInfo == null)
                {
                    bodyInfo = new EPPlusConfigBodyInfo()
                    {
                        Nth = i + 1,
                        Option = new EPPlusConfigBodyOption(),
                    };
                    config.Body.InfoList.Add(bodyInfo);
                }

                bodyInfo.Option.SummaryMapperExcel = dictSummeryList[i];
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
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                var colName = dt.Columns[i].ColumnName;
                if (!dict.ContainsKey(colName))
                {
                    dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
                }
                else
                {
                    throw new Exception(nameof(SetConfigSourceHead) + "方法异常");
                }
            }

            var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell>();
            foreach (var item in dict)
            {
                fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key, FillValue = dict.Values });
            }

            configSource.Head = new EPPlusConfigSourceHead() { CellsInfoList = fixedCellsInfoList };
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
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                var colName = dt.Columns[i].ColumnName;
                if (!dict.ContainsKey(colName))
                {
                    dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
                }
                else
                {
                    throw new Exception(nameof(SetConfigSourceFoot) + "方法异常");
                }
            }

            var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell>();
            foreach (var item in dict)
            {
                fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key, FillValue = dict.Values });
            }

            configSource.Foot = new EPPlusConfigSourceFoot { CellsInfoList = fixedCellsInfoList };
        }

        #endregion

    }
}
