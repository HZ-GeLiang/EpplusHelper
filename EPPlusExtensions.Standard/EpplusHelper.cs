using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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
using OfficeOpenXml.Style;

namespace EPPlusExtensions
{
    public class EPPlusHelper
    {

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
        /// 删除母版页
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
        /// 删除母版页
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
            config.WorkSheetDefault?.Invoke(worksheet);
            FillData(config, configSource, worksheet);
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
            ExcelWorksheet worksheet = GetExcelWorksheet(excelPackage, destWorkSheetIndex, workSheetNewName);
            config.WorkSheetDefault?.Invoke(worksheet);
            FillData(config, configSource, worksheet);
        }

        /// <summary>
        /// 往目标sheet中填充数据
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        private static void FillData(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            FillData_Head(config, configSource, worksheet);
            var sheetBodyAddRowCount = FillData_Body(config, configSource, worksheet);
            FillData_Foot(config, configSource, worksheet, sheetBodyAddRowCount);
        }


        #endregion

        #region 私有方法

        private static void FillData_Head(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            //填充head
            foreach (var item in config.SheetHeadMapperExcel)
            {
                if (configSource.SheetHead.Keys.Count == 0) //excel中有配置head,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
                {
                    break;
                }

                //var cellpoint = new ExcelCellPoint(item.Key);
                string colMapperName = item.Value;
                object val = configSource.SheetHead[item.Value];
                //ExcelRange cells = worksheet.Cells[cellpoint.Row , cellpoint.Col];
                ExcelRange cells = worksheet.Cells[item.Key];

                if (config.SheetHeadCellCustomSetValue != null)
                {
                    config.SheetHeadCellCustomSetValue.Invoke(colMapperName, val, cells);
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
            int sheetBodyDeleteRowCount = 0; //sheet body 中删除了多少行(只含配置的行,对于FillData()内的删除行则不包括在内).  
            var sheetBodyAddRowCount = 0; //新增了几行 (统计sheet body 在原有的模版上新增了多少行)
            foreach (var nth in config.SheetBodyMapperExcel) //body的第N个配置
            {
                #region get dataTable
                DataTable datatable;
                if (!configSource.SheetBody.ContainsKey(nth.Key)) //如果没有数据源中没有excle中配置
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
                    datatable = configSource.SheetBody[nth.Key]; //body的第N个配置的数据源
                }

                #endregion

                #region When dataTable is empty
                if (datatable == null || datatable.Rows.Count <= 0) //数据源为null或为空
                {
                    //throw new ArgumentNullException($"configSource.SheetBody[{nth.Key}]没有可读取的数据");

                    if (!config.DeleteFillDateStartLineWhenDataSourceEmpty || nth.Value.Keys.Count <= 0)
                    {
                        continue; //跳过本次fillDate的循环
                    }

                    #region DeleteFillDateStartLine

                    foreach (var r1c1 in nth.Value.Keys) //只遍历一次
                    {
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
                }
                #endregion

                int currentLoopAddLines = 0;
                var deleteLastSpaceLine = false; //是否删除最后一空白行(可能有多行组成的)
                int lastSpaceLineInterval = 0; //表示最后一空白行由多少行组成,默认为0
                int lastSpaceLineRowNumber = 0; //表示最后一行的行号是多少
                int tempLine = config.SheetBodyMapperExceltemplateLine.ContainsKey(nth.Key)
                    ? config.SheetBodyMapperExceltemplateLine[nth.Key]
                    : 1; //获得第N个配置中excel模版提供了多少行,默认1行
                var hasMergeCell = nth.Value.Keys.ToList().Find(a => a.Contains(":")) != null;
                Dictionary<string, FillDataColums> fillDataColumsStat = null;//Datatable 的列的使用情况

                if (hasMergeCell)
                {
                    //注:进入这里的条件是单元格必须是多行合并的,如果是同行多列合并的单元格,最后生成的excel会有问题,打开时会提示修复(修复完成后内容是正确的(不保证,因为我测试的几个内容是正确的))
                    List<ExcelCellRange> cellRange = nth.Value.Keys.Select(r1C1 => new ExcelCellRange(r1C1)).ToList();

                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        DataRow row = datatable.Rows[i];
                        int destRow;
                        int maxIntervalRow = (from c in cellRange select c.IntervalRow).Max();

                        if (nth.Key == 1)
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
                                worksheet.InsertRow(destRow, maxIntervalRow + 1,
                                    destRow + maxIntervalRow + 1); //新增N行,注意,此时新增行的高度是有问题的
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
                        for (int j = 0; j < cellRange.Count; j++)
                        {
                            #region 赋值
                            string colMapperName = nth.Value[cellRange[j].Range];
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

                            if (config.SheetBodyCellCustomSetValue.ContainsKey(nth.Key) && config.SheetBodyCellCustomSetValue[nth.Key] != null)
                            {
                                config.SheetBodyCellCustomSetValue[nth.Key].Invoke(colMapperName, val, cells);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }
                            #endregion

                            #region 同步数据源
                            if (j == cellRange.Count - 1) //如果一行循环到了最后一列
                            {
                                if (!configSource.SheetBodyFillModel.ContainsKey(nth.Key))
                                {
                                    continue;
                                }
                                var fillModel = configSource.SheetBodyFillModel[nth.Key];
                                if (fillModel == null)
                                {
                                    continue;
                                }
                                if (fillModel.FillDataMethodOption == SheetBodyFillDataMethodOption.Default)
                                {
                                    continue;
                                }

                                if (fillModel.FillDataMethodOption == SheetBodyFillDataMethodOption.SynchronizationDataSource)
                                {
                                    var isFillData_Title = fillModel.SynchronizationDataSource.NeedTitle && i == 0;
                                    var isFillData_Body = fillModel.SynchronizationDataSource.NeedBody;
                                    if ((isFillData_Title) || isFillData_Body)
                                    {
                                        if (fillDataColumsStat == null)
                                        {
                                            fillDataColumsStat = InitFlilDataColumsStat(datatable, nth, fillModel);
                                        }

                                        if (isFillData_Title)
                                        {
                                            var eachCount = 0;
                                            var config_firstCell_col = new ExcelCellPoint(nth.Value.First().Key).Col;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {
                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                var extensionDestCol_title = config_firstCell_col + nth.Value.Count + eachCount;
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
                    var startCellPointLine = nth.Value.Keys.Select(a => new ExcelCellPoint(a)).ToList(); // 将配置的值 转换成 ExcelCellPoint

                    #region 等价于上面的写法

                    //List<ExcelCellPoint> startCellPointLine = new List<ExcelCellPoint>();
                    //foreach (var r1c1 in nth.Value.Keys) //将配置的值 转换成 ExcelCellPoint
                    //{
                    //    startCellPointLine.Add(R1C1ToExcelCellPoint(r1c1));
                    //    //int row = Convert.ToInt32(RegexHelper.GetLastNumber(item.Key));
                    //    //string col = RegexHelper.GetFirstStringByReg(item.Key, "[A-Za-z]+");
                    //    //startCellPointLine.Add(new ExcelCellPoint(row, col, item.Key));
                    //}

                    #endregion

                    for (int i = 0; i < datatable.Rows.Count; i++) //遍历数据源,往excel中填充数据
                    {
                        DataRow row = datatable.Rows[i];
                        int destRow;
                        if (nth.Key == 1)
                        {
                            destRow = sheetBodyAddRowCount > 0
                                ? startCellPointLine[0].Row + i - sheetBodyDeleteRowCount
                                : startCellPointLine[0].Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }
                        else
                        {
                            destRow = currentLoopAddLines > 0
                                ? startCellPointLine[0].Row + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                                : startCellPointLine[0].Row + i + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }

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
                                //必须先新增,在赋值(若先赋值后新增,会造成赋值后的行被新增行覆盖).
                                //1.新增一行,在destRow 前 插入 1行
                                worksheet.InsertRow(destRow, 1, destRow + 1);
                                //copyStylesFromRow参数不会把合并的单元格也弄过来(即,这个参数的功能不是格式刷)
                                //worksheet.InsertRow(destRow, 1);//注,这行代码与上一行代码的作用是一样的,因为我下面用了Copy.

                                //2.复制样式(含修正)
                                //然后把原本的destRow的样式格式化到新增行中.注意:copy 会把copy行的文本也复制出来.
                                //这里可以说是一个潜在的隐患bug把.因为和我的本意不一样.主要是我不知道要怎么写,只找到一个copy方法,而且copy方法也能帮我解决掉 同一行的 单元格合并问题
                                string copyRowSource = (destRow + 1) + ":" + (destRow + 1); //7:7表示第7行,Copy中的Dest行
                                string copyRowDest = (destRow) + ":" + (destRow);
                                worksheet.Cells[copyRowSource].Copy(worksheet.Cells[copyRowDest]);
                                //不要用[row,col]索引器,[row,col]表示某单元格.注意:copy会把source行的除了height(觉得是一个bug)以外的全部复制一行出来
                                worksheet.Row(destRow).Height = worksheet.Row(destRow + 1).Height; //修正height

                                sheetBodyAddRowCount++;
                                currentLoopAddLines++;
                            }
                        }

                        //3.赋值.
                        //注:遍历时变量 j 的终止条件不能是 datatable.Rows.Count. 因为datatable可能会包含多余的字段信息,与 配置信息列的个数不一致.
                        for (int j = 0; j < startCellPointLine.Count; j++)
                        {
                            #region 赋值
                            //worksheet.Cells[destRow, destCol].Value = row[j];
                            string colMapperName = nth.Value[startCellPointLine[j].R1C1];
                            object val = row[colMapperName];
                            int destCol = startCellPointLine[j].Col;
                            ExcelRange cells = worksheet.Cells[destRow, destCol];
                            if (config.SheetBodyCellCustomSetValue.ContainsKey(nth.Key) && config.SheetBodyCellCustomSetValue[nth.Key] != null)
                            {
                                config.SheetBodyCellCustomSetValue[nth.Key]?.Invoke(colMapperName, val, cells);
                            }
                            else
                            {
                                SetWorksheetCellsValue(config, cells, val, colMapperName);
                            }
                            #endregion

                            #region 同步数据源
                            if (j == startCellPointLine.Count - 1) //如果一行循环到了最后一列
                            {
                                if (!configSource.SheetBodyFillModel.ContainsKey(nth.Key))
                                {
                                    continue;
                                }
                                var fillModel = configSource.SheetBodyFillModel[nth.Key];
                                if (fillModel == null)
                                {
                                    continue;
                                }
                                if (fillModel.FillDataMethodOption == SheetBodyFillDataMethodOption.Default)
                                {
                                    continue;
                                }

                                if (fillModel.FillDataMethodOption == SheetBodyFillDataMethodOption.SynchronizationDataSource)
                                {
                                    var isFillData_Title = fillModel.SynchronizationDataSource.NeedTitle && i == 0;
                                    var isFillData_Body = fillModel.SynchronizationDataSource.NeedBody;
                                    if ((isFillData_Title) || isFillData_Body)
                                    {
                                        if (fillDataColumsStat == null)
                                        {
                                            fillDataColumsStat = InitFlilDataColumsStat(datatable, nth, fillModel);
                                        }

                                        if (isFillData_Title)
                                        {
                                            var eachCount = 0;
                                            var config_firstCell_col = new ExcelCellPoint(nth.Value.First().Key).Col;
                                            foreach (var item in fillDataColumsStat.Values)
                                            {
                                                if (item.State != FillDataColumsState.WillUse) continue;
                                                var extensionDestCol_title = config_firstCell_col + nth.Value.Count + eachCount;
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
                                                var extensionDestCol = startCellPointLine[j].Col + 1 + eachCount;

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

                //已经修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白 
                if (deleteLastSpaceLine)
                {
                    worksheet.DeleteRow(lastSpaceLineRowNumber, lastSpaceLineInterval, true);
                    sheetBodyAddRowCount -= lastSpaceLineInterval;
                }

                FillData_Body_Summary(config, configSource, worksheet, nth, sheetBodyAddRowCount);
            }

            return sheetBodyAddRowCount;
        }

        /// <summary>
        ///  获得Database数据源的所有的列的使用状态
        /// </summary>
        /// <param name="fillDataColumsStat"></param>
        /// <param name="datatable"></param>
        /// <param name="nth"></param>
        /// <param name="fillModel"></param>
        /// <returns></returns>
        private static Dictionary<string, FillDataColums> InitFlilDataColumsStat(DataTable datatable, KeyValuePair<int, Dictionary<string, string>> nth, SheetBodyFillDataMethod fillModel)
        {
            var fillDataColumsStat = new Dictionary<string, FillDataColums>();
            foreach (DataColumn column in datatable.Columns)
            {
                fillDataColumsStat.Add(column.ColumnName, new FillDataColums()
                {
                    ColumName = column.ColumnName,
                    State = FillDataColumsState.Unchanged
                });
            }

            foreach (var item in nth.Value)
            {
                fillDataColumsStat[item.Value].State = FillDataColumsState.Used;
            }

            var isEmptyInclude = string.IsNullOrEmpty(fillModel.SynchronizationDataSource.Include);
            var isEmptyExclude = string.IsNullOrEmpty(fillModel.SynchronizationDataSource.Exclude);
            if (isEmptyInclude != isEmptyExclude) //只能有一个值有效
            {
                if (!isEmptyInclude)
                {
                    Modify_DataColumnsIsUsedStat(fillDataColumsStat, fillModel.SynchronizationDataSource.Include, true);
                }

                if (!isEmptyExclude)
                {
                    Modify_DataColumnsIsUsedStat(fillDataColumsStat, fillModel.SynchronizationDataSource.Exclude, false);
                }
            }

            return fillDataColumsStat;
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

        private static void FillData_Body_Summary(EPPlusConfig config, EPPlusConfigSource configSource,
            ExcelWorksheet worksheet, KeyValuePair<int, Dictionary<string, string>> nth, int sheetBodyAddRowCount)
        {
            if (!config.SheetBodySummaryMapperExcel.Keys.Contains(nth.Key)) return;

            foreach (var item in config.SheetBodySummaryMapperExcel[nth.Key]) //填充第N个配置的一些零散的单元格的值(譬如汇总信息等)
            {
                var cellpoint = new ExcelCellPoint(item.Key);
                // worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col].Value = configSource.SheetBodySummary[nth.Key][item.Value];
                //item.Key -> A24 , item.Value -> 平均值
                string colMapperName = item.Value;
                object val = configSource.SheetBodySummary[nth.Key][item.Value];
                ExcelRange cells = worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col];

                if (config.SheetBodySummaryCellCustomSetValue.ContainsKey(nth.Key) &&
                    config.SheetBodySummaryCellCustomSetValue[nth.Key] != null)
                {
                    config.SheetBodySummaryCellCustomSetValue[nth.Key].Invoke(colMapperName, val, cells);
                }
                else
                {
                    SetWorksheetCellsValue(config, cells, val, colMapperName);
                }
            }
        }

        private static void FillData_Foot(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet,
            int sheetBodyAddRowCount)
        {
            //填充foot
            foreach (var item in config.SheetFootMapperExcel)
            {
                if (configSource.SheetFoot.Keys.Count == 0) //excel中有配置foot,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
                {
                    break;
                }

                //worksheet.Cells["A1"].Value = "名称";//直接指定单元格进行赋值
                //var cellpoint = R1C1ToExcelCellPoint(item.Key);
                var cellpoint = new ExcelCellPoint(item.Key);
                // worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col].Value = configSource.SheetFoot[item.Value];
                string colMapperName = item.Value;
                object val = configSource.SheetFoot[item.Value];
                ExcelRange cells = worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col];
                if (config.SheetFootCellCustomSetValue != null)
                {
                    config.SheetFootCellCustomSetValue.Invoke(colMapperName, val, cells);
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
        ///  从Excel 中获得符合C# 类属性定义的列名集合
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">列名在Excel的第几行</param>
        /// <returns></returns>
        private static List<ExcelCellInfo> GetExcelColumnOfModel(ExcelWorksheet ws, int row, int colStart, int? colEnd)
        {
            if (colEnd == null) colEnd = EPPlusConfig.MaxCol07;
            var list = new List<ExcelCellInfo>();
            for (int col = colStart; col < colEnd; col++)
            {
                string colName = ws.Cells[row, col].Text;
                if (string.IsNullOrEmpty(colName)) break;
                colName = ExtractName(colName);
                if (string.IsNullOrEmpty(colName)) break;
                list.Add(new ExcelCellInfo()
                {
                    WorkSheet = ws,
                    Value = colName,
                    ExcelCellPoint = new ExcelCellPoint(row, col)
                });
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
        /// excel的所有列均在model对象的属性中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="colNameList"></param>
        /// <param name="modelCheckMsg"></param>
        /// <returns></returns>
        private static bool IsAllExcelColumnExistsModel<T>(List<ExcelCellInfo> colNameList, out string modelCheckMsg) where T : class, new()
        {
            StringBuilder sb = new StringBuilder();
            Type type = typeof(T);
            foreach (var item in colNameList)
            {
                PropertyInfo pInfo = type.GetProperty(item.Value.ToString());
                if (pInfo == null)
                {
                    sb.AppendLine($"WorkSheet:'{item.WorkSheet.Name}' 的'{item.ExcelCellPoint.R1C1}'值'{item.Value}'在'{type.FullName}'类型中没有定义该属性");
                }
            }
            modelCheckMsg = sb.ToString();
            return sb.Length <= 0;
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
            object[] validAttrs = ReflectionHelper.GetAttributeForProperty<T, System.ComponentModel.DataAnnotations.ValidationAttribute>(pInfo.Name, true);
            if (validAttrs != null && validAttrs.Length > 0)
            {
                //同一个特性的的属性值肯定是一样的,所以可以优化;
                //ArrayList objArr = null;
                Object[] objArr = null;//第二次优化
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
                    var IsValid = (bool)methodIsValid.Invoke(validAttr, objArr);

                    #endregion

                    if (!IsValid)
                    {
                        var msg = $@"'{model.GetType().FullName}'类型的'{pInfo.Name}'属性验证未通过:'{((System.ComponentModel.DataAnnotations.ValidationAttribute)validAttr).ErrorMessage}'";
                        throw new ArgumentException(msg);

                    }
                }
            }
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
                    throw new ArgumentException("无效的日期", nameof(pInfo.Name));
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
                    throw new ArgumentException("无效的数字", nameof(pInfo.Name));
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
                    throw new ArgumentException("无效的数字", nameof(pInfo.Name));
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
                if (!Int16.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", nameof(pInfo.Name));
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
                if (!Int16.TryParse(value, out var result))
                {
                    throw new ArgumentException("无效的数字", nameof(pInfo.Name));
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
                    throw new ArgumentException($@"无效的{pInfo_PropertyType.FullName}枚举值", nameof(pInfo.Name));
                }
                value = ExtractName(value);
                TryThrowExceptionForEnum(pInfo, model, value, pInfo_PropertyType, pInfo_PropertyType);
                var enumValue = Enum.Parse(pInfo_PropertyType, value);
                pInfo.SetValue(model, enumValue);
                return;
            }
            #endregion

            throw new System.Exception("未考虑到的情况!!!请完善程序");
        }

        private static void TryThrowExceptionForEnum<T>(PropertyInfo pInfo, T model, string value, Type enumType, Type pInfoType) where T : class, new()
        {
            var isDefined = Enum.IsDefined(enumType, value);
            if (isDefined)
            {
                return;
            }
            var attrs = ReflectionHelper.GetAttributeForProperty<EnumUndefinedAttribute>(pInfo.DeclaringType, pInfo.Name);
            if (attrs.Length != 1)
            {
                throw new System.ArgumentException($"Value值:'{value}'在枚举值:'{pInfoType.FullName}'中未定义,请检查!!!");
            }

            var attr = (EnumUndefinedAttribute)attrs[0];
            if (attr.Args == null || attr.Args.Length <= 0)
            {
                throw new System.ArgumentException($"Value值:'{value}'在枚举值:'{pInfoType.FullName}'中未定义,请检查!!!");
            }

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

        public static EPPlusConfig GetEmptyConfig() => new EPPlusConfig();

        public static EPPlusConfigSource GetEmptyConfigSource() => new EPPlusConfigSource();

        #endregion

        #region GetList<T>

        /// <summary>
        /// 只能是最普通的excel.(每个单元格都是未合并的,第一行是列名,数据从第一列开始填充的那种.)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowIndex">数据起始行(不含列名),从1开始</param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex) where T : class, new()
        {
            return GetList<T>(new GetExcelListArgs<T>()
            {
                ws = ws,
                rowIndex_Data = rowIndex,
                EveryCellPrefix = "",
                EveryCellReplace = null,
                rowIndex_DataName = rowIndex - 1,
                UseEveryCellReplace = true,
                HavingFilter = null,
                WhereFilter = null,
                ReadCellValueOption = ReadCellValueOption.Trim,
            });
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
            return GetList<T>(new GetExcelListArgs<T>()
            {
                ws = ws,
                rowIndex_Data = rowIndex,
                EveryCellPrefix = everyCellPrefix,
                EveryCellReplace =
                    everyCellReplaceOldValue == null || everyCellReplaceNewValue == null
                    ? null
                    : new Dictionary<string, string> { { everyCellReplaceOldValue, everyCellReplaceNewValue } },
                rowIndex_DataName = rowIndex - 1,
                UseEveryCellReplace = true,
                HavingFilter = null,
                WhereFilter = null,
                ReadCellValueOption = ReadCellValueOption.Trim,
            });
        }

        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellPrefix, Dictionary<string, string> everyCellReplace) where T : class, new()
        {
            return GetList<T>(new GetExcelListArgs<T>()
            {
                ws = ws,
                rowIndex_Data = rowIndex,
                EveryCellPrefix = everyCellPrefix,
                EveryCellReplace = everyCellReplace,
                rowIndex_DataName = rowIndex - 1,
                UseEveryCellReplace = true,
                HavingFilter = null,
                WhereFilter = null,
                ReadCellValueOption = ReadCellValueOption.Trim,
            });
        }

        public static List<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.rowIndex_Data;
            string everyCellPrefix = args.EveryCellPrefix;
            int dataNameRowIndex = args.rowIndex_DataName;
            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplace == null
                ? GetExcelListArgs<T>.EveryCellReplaceDefault
                : args.EveryCellReplace;
            var havingFilter = args.HavingFilter;
            var whereFilter = args.WhereFilter;
            var readCellValueOption = args.ReadCellValueOption;

            if (rowIndex == default(int) || dataNameRowIndex == default(int))
            {
                throw new ArgumentException("请初始化数据");
            }

            List<T> list = new List<T>();

            var colNameList = GetExcelColumnOfModel(ws, dataNameRowIndex, 1, EPPlusConfig.MaxCol07);
            var dictColName = colNameList.ToDictionary(item => item.ExcelCellPoint.Col, item => item.Value.ToString());

            string modelCheckMsg;
            if (!IsAllExcelColumnExistsModel<T>(colNameList, out modelCheckMsg)) throw new ExcelColumnNotExistsWithModelException(modelCheckMsg);

            bool canEveryCellReplace = everyCellReplace != null;

            for (int row = rowIndex; row <= EPPlusConfig.MaxRow07; row++)
            {
                if (string.IsNullOrEmpty(ws.Cells[row, 1].Text))//每一行的第一列为空
                {
                    if (row == rowIndex)
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    else
                    {
                        break; //读取模版结束
                    }
                }

                if (ws.Cells[row, 1].Merge)
                {
                    throw new System.ArgumentException($@"数据的每一行的首列不能有合并单元格,当前行是第{row}行");
                }

                Type type = typeof(T);
                var ctor = type.GetConstructor(new Type[] { });
                if (ctor == null) throw new ArgumentException($"通过反射无法得到{type.FullName}的一个无构造参数的构造器:");
                T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转

                for (int col = 1; col < EPPlusConfig.MaxCol07; col++)
                {
                    if (!dictColName.ContainsKey(col)) break;
                    string colName = dictColName[col];
                    if (string.IsNullOrEmpty(colName)) break;
                    PropertyInfo pInfo = type.GetProperty(colName);
                    if (pInfo == null)
                    {
                        throw new ArgumentException($@"Type:{type} 的 property {colName} 未找到");
                    }
                    string value;
                    if (ws.Cells[row, col].Merge)
                    {
                        value = EPPlusHelper.GetMegerCellText(ws, row, col);
                    }
                    else
                    {
                        value = ws.Cells[row, col].Text;
                    }

                    if (value == null || value.Length <= 0)
                    {
                        GetList_SetModelValue(pInfo, model, value);
                        continue;
                    }

                    if (everyCellPrefix?.Length > 0)
                    {
                        var indexof = value.IndexOf(everyCellPrefix);
                        if (indexof == -1)
                        {
                            throw new System.ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + everyCellPrefix + "'开头的");
                        }
                        value = value.RemovePrefix(everyCellPrefix);
                    }
                    if (canEveryCellReplace)
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
                    switch (args.ReadCellValueOption)
                    {
                        case ReadCellValueOption.None:
                            break;
                        case ReadCellValueOption.Trim:
                            value = value.Trim();
                            break;
                        case ReadCellValueOption.MergeLine:
                            value = value.MergeLines();
                            break;
                        case ReadCellValueOption.MergeLineAndTrim:
                            value = value.Trim().MergeLines().Trim();
                            break;
                        default:
                            throw new System.Exception("未指定读取单元格值时的操作方式");
                    }
                    GetList_SetModelValue(pInfo, model, value);
                }

                if (whereFilter == null || whereFilter.Invoke(model))
                {
                    list.Add(model);
                }
            }
            return havingFilter == null ? list : list.Where(item => havingFilter.Invoke(item)).ToList();
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
        public static string GetMegerCellText(ExcelWorksheet ws, int row, int col)
        {
            string range = ws.MergedCells[row, col];
            if (range == null) return GetCellText(ws, row, col);
            var ea = new ExcelAddress(range).Start;
            return ws.Cells[ea.Row, ea.Column].Text;
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
            var cell = ws.Cells[row, col];
            //if (cell.Merge) throw new Exception("没遇到过这个情况的");
            // return cell.Text; //这个没有科学计数法  注:Text是Excel显示的值,Value是实际值.
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
                else
                {
                    throw;
                }
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
                            ExcelCellPoint = new ExcelCellPoint(i + 1, j + 1),
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
        /// <returns>是否有进行科学技术法的cell转换.true:是,fale:否</returns>
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
        /// <returns>是否有进行科学技术法的cell转换.true:是,fale:否</returns>
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
            SetDefaultConfigFromExcel(excelPackage, config, sheet);
        }

        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, string workSheetName)
        {
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));
            var sheet = GetExcelWorksheet(excelPackage, workSheetName);
            SetDefaultConfigFromExcel(excelPackage, config, sheet);
        }
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, ExcelWorksheet sheet)
        {
            //让 sheet.Cells.Value 强制从A1单元格开始
            //遇到问题描述:创建一个exccel,在C7,C8,C9,10单元格写入一些字符串, sheet.Cells.Value 是object[4,3]的数组, 但我要的是object[10,3]的数组
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value == null)
            {
                cellA1.Value = null;
            }

            SetConfigHeadFromExcel(excelPackage, config, sheet);
            SetConfigBodyFromExcel(excelPackage, config, sheet);
            SetConfigFootFromExcel(excelPackage, config, sheet);
        }

        /// <summary>
        /// 设置sheetHead配置
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetConfigHeadFromExcel(ExcelPackage excelPackage, EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (cellStr.StartsWith("$tb")) //说明$th的配置已经结束了
                    {
                        break;
                    }

                    if (!cellStr.StartsWith("$th")) continue;

                    //{"G6", "公司名称"},
                    string key = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);

                    string val = Regex.Replace(cellStr, "^[$]th", ""); //$需要
                    if (dict.ContainsValue(val))
                    {
                        throw new ArgumentException($"Excel文件中的$th部分配置了相同的项:{val}");
                    }

                    dict.Add(key.Trim(), val.Trim());
                    //arr[i,j] = "";//把当前单元格值清空
                    sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉
                }
            }
            config.SheetHeadMapperExcel = dict;
        }

        /// <summary>
        /// 设置sheetBody配置
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetConfigBodyFromExcel(ExcelPackage excelPackage, EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            var dictList = new List<Dictionary<string, string>>();
            var dictSummeryList = new List<Dictionary<string, string>>();
            var sheetMergedCellsList = sheet.MergedCells.ToList();

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
                    string key = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);

                    string nthStr = RegexHelper.GetFirstNumber(cellStr);
                    int nth = Convert.ToInt32(nthStr);
                    if (cellStr.StartsWith("$tbs")) //模版摘要/汇总等信息单元格
                    {
                        string val = Regex.Replace(cellStr, "^[$]tbs" + nthStr, ""); //$需要转义
                        if (dictSummeryList.Count < nth)
                        {
                            dictSummeryList.Add(new Dictionary<string, string>());
                        }
                        if (dictSummeryList[nth - 1].ContainsValue(val))
                        {
                            throw new ArgumentException($"Excel文件中的$tbs{nth}部分配置了相同的项:{val}");
                        }
                        dictSummeryList[nth - 1].Add(key.Trim(), val.Trim());
                    }
                    else if (cellStr.StartsWith($"$tb{nthStr}$")) //模版提供了多少行,若没有配置,在调用FillData()时默认提供1行
                    {
                        string valStr = Regex.Replace(cellStr, $@"^[$]tb{nth}[$]", ""); //$需要转义
                        var val = string.Compare(valStr, "max", true) == 0 //$tb1$max这种配置的
                            ? EPPlusConfig.MaxRow07 - i
                            : Convert.ToInt32(valStr);
                        if (config.SheetBodyMapperExceltemplateLine.ContainsKey(val))
                        {
                            throw new ArgumentException($"Excel文件中重复配置了tb{nthStr}的行数");
                        }
                        config.SheetBodyMapperExceltemplateLine.Add(nth, val);
                    }
                    else //StartsWith($"$tb{nthStr}")
                    {
                        string val = Regex.Replace(cellStr, "^[$]tb" + nthStr, ""); //$需要转义

                        if (dictList.Count < nth)
                        {
                            dictList.Add(new Dictionary<string, string>());
                        }
                        if (dictList[nth - 1].ContainsValue(val))
                        {
                            throw new ArgumentException($"Excel文件中的$tb{nth}部分配置了相同的项:{val}");
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
                            var newKey = sheetMergedCellsList.Find(a => a.Contains(key));
                            if (newKey == null)
                            {
                                //描述出现null的情况
                                /*
                                 * F10 G10
                                 * F11 G11
                                 * F12 G12
                                 * 这些单元格被合并为一个单元格,即用F10:G12来描述
                                 * 此时,配置单元格读取应该是F10,G10将不会被读取,
                                 * 直到上面为止,都是正确的,但是,偏偏有一个神一样的工具,
                                 * 当excel模版出现不规范操作(Excel一眼看上去将没有问题),G10单元格被读取出来后,那么在sheetMergedCellsList中肯定找不到
                                 * 然后下面一行代码就抛出未将对象引用设置到对象的实例异常
                                 * 该操作是:B10, D10, F10, G10单元格均有配置项,B10:C12进行单元格合并,然后用格式刷,对D10:E12, F10:G10进行格式化
                                 * 结果就是G10的配置项将会被隐藏
                                 * 如果手动的合并F10:G10,Excel将会alert此操作会仅保留左上角的值
                                 */
                                throw new Exception($"excel的单元格{key}存在配置问题,请检查");
                            }
                            var cells = newKey.Split(':');
                            var dictListKey = RegexHelper.GetFirstNumber(cells[0]) == RegexHelper.GetFirstNumber(cells[1]) ? key : newKey;
                            dictList[nth - 1].Add(dictListKey, val);
                        }
                        else
                        {
                            dictList[nth - 1].Add(key.Trim(), val.Trim());
                        }
                    }

                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉
                    sheet.Cells[i + 1, j + 1].Value = null;//修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白
                                                           //如果有用 sheet.Cells[i + 1, j + 1].Value = "" 代码 ,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                }

            }
            for (int i = 0; i < dictList.Count; i++)
            {
                config.SheetBodyMapperExcel.Add(i + 1, dictList[i]); //索引从1开始,所以这边要+1
            }
            for (int i = 0; i < dictSummeryList.Count; i++)
            {
                config.SheetBodySummaryMapperExcel.Add(i + 1, dictSummeryList[i]); //索引从1开始,所以这边要+1
            }

        }

        /// <summary>
        /// 设置sheetFoot配置
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static void SetConfigFootFromExcel(ExcelPackage excelPackage, EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            var dict = new Dictionary<string, string>();
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (!cellStr.StartsWith("$tf")) continue;

                    // {"G6", "公司名称"},
                    string key = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);
                    string val = Regex.Replace(cellStr, "^[$]tf", ""); //$需要转义
                    if (dict.ContainsValue(val))
                    {
                        throw new ArgumentException($"Excel文件中的$tf部分配置了相同的项:{val}");
                    }

                    dict.Add(key.Trim(), val.Trim());
                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉
                    sheet.Cells[i + 1, j + 1].Value = null; //统一用 null 来清空单元格
                }
            }
            config.SheetFootMapperExcel = dict;
        }


        #endregion

        #region 设置Head与foot配置的数据源

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
                dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
            }
            configSource.SheetHead = dict;
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
                dict.Add(colName, dr[i] == DBNull.Value || dr[i] == null ? "" : dr[i].ToString());
            }
            configSource.SheetFoot = dict;
        }

        #endregion

        #region 对单元格样式进行 Get Set

        /// <summary>
        ///  获取Cell样式
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static EPPlusCellStyle GetCellStyle(ExcelRange cell)
        {
            EPPlusCellStyle cellStyle = new EPPlusCellStyle();
            cellStyle.HorizontalAlignment = cell.Style.HorizontalAlignment;
            cellStyle.VerticalAlignment = cell.Style.VerticalAlignment;
            cellStyle.WrapText = cell.Style.WrapText;
            cellStyle.FontBold = cell.Style.Font.Bold;
            cellStyle.FontColor = string.IsNullOrEmpty(cell.Style.Font.Color.Rgb)
                ? Color.Black
                : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Font.Color.Rgb);
            cellStyle.FontName = cell.Style.Font.Name;
            cellStyle.FontSize = cell.Style.Font.Size;
            cellStyle.BackgroundColor = string.IsNullOrEmpty(cell.Style.Fill.BackgroundColor.Rgb)
                ? Color.Black
                : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Fill.BackgroundColor.Rgb);
            cellStyle.ShrinkToFit = cell.Style.ShrinkToFit;

            return cellStyle;
        }

        /// <summary>
        /// 设置Cell样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        public static void SetCellStyle(ExcelRange cell, EPPlusCellStyle style)
        {
            cell.Style.HorizontalAlignment = style.HorizontalAlignment;
            cell.Style.VerticalAlignment = style.VerticalAlignment;
            cell.Style.WrapText = style.WrapText;
            cell.Style.Font.Bold = style.FontBold;
            cell.Style.Font.Color.SetColor(style.FontColor);
            if (!string.IsNullOrEmpty(style.FontName))
            {
                cell.Style.Font.Name = style.FontName;
            }
            cell.Style.Font.Size = style.FontSize;
            cell.Style.Fill.PatternType = style.PatternType;
            cell.Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
            cell.Style.ShrinkToFit = style.ShrinkToFit;
        }

        #endregion

        #region 一些默认的sql语句

        /// <summary>
        /// 获得树形表结构的最深的层级数的Sql语句
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="rootItemWhere">root(根)数据的where条件,即根据表名获得root(根)数据的条件是什么</param>
        public static string GetTreeTableMaxLevelSql(string tblName, string rootItemWhere, string idFiledName = "Id", string parentIdName = "ParentId")
        {
            string sql = $@"with cte as( 
            SELECT {idFiledName} ,  1 as level FROM {tblName} WHERE {rootItemWhere}
            union all 
            SELECT {tblName}.{idFiledName}, cte.level+1 as level from cte, {tblName}  where cte.{idFiledName} = {tblName}.{parentIdName} 
            )
            SELECT ISNULL(MAX(cte.level),0) FROM  cte";
            return sql;
        }

        /// <summary>
        /// 原本的树形表结构是没有Level字段的,通过该方法可以生成level字段
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="rootItemWhere"></param>
        /// <param name="nameFieldName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="otherFiledName"></param>
        /// <returns></returns>
        public static string GetTreeTableIncludeLevelFieldSql(string tblName, string rootItemWhere, string nameFieldName = "Name", string idFiledName = "Id", string parentIdName = "ParentId", params string[] otherFiledName)
        {
            string comma = " ,";
            string dot = ".";
            StringBuilder sb1 = new StringBuilder(); //定位成员的字段
            sb1.Append(idFiledName).Append(comma)
                .Append(nameFieldName).Append(comma)
                .Append(parentIdName);
            StringBuilder sb2 = new StringBuilder(); //递归成员的字段
            sb2.Append(tblName).Append(dot).Append(idFiledName).Append(comma)
                .Append(tblName).Append(dot).Append(nameFieldName).Append(comma)
                .Append(tblName).Append(dot).Append(parentIdName);

            if (otherFiledName != null && otherFiledName.Length > 0)
            {
                foreach (var item in otherFiledName)
                {
                    sb1.Append(item).Append(comma);
                    sb2.Append(tblName).Append(dot).Append(item).Append(comma);
                }
                sb1.RemoveLastChar(comma.Length);
                sb2.RemoveLastChar(comma.Length);
            }

            string sql = $@"with cte as( 
            SELECT {sb1} , 1 as Level FROM {tblName} WHERE {rootItemWhere}
            union all 
            SELECT {sb2} , cte.Level+1 as Level from cte, {tblName}  
                where cte.{idFiledName} = {tblName}.{parentIdName} 
            )
            SELECT {sb1} , Level FROM  cte
            ORDER BY cte.Level";
            return sql;
        }

        /// <summary>
        ///  根据 id, Name , parentId 3个字段生成额外字段Depth 和 用于报表排序的Sort字段
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="rootItemWhere"></param>
        /// <param name="nameFieldName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="eachSortFieldLength">每个Depth的长度,默认2. </param>
        /// <param name="reportSortFileTotallength">报表排序字段的总长度,默认为12如果真的要设置,level * Max(Len(主键))</param>
        /// <param name="rearChat">报表排序字段 / 每个Depth字段 小于 指定长度时填充的字符是什么</param>
        /// <param name="otherFiledName"></param>
        /// <returns></returns>
        public static string GetTreeTableReportSql(string tblName, string rootItemWhere, string nameFieldName = "Name", string idFiledName = "Id", string parentIdName = "ParentId", int eachSortFieldLength = 2, int reportSortFileTotallength = 12, char rearChat = ' ', params string[] otherFiledName)
        {
            //该方法基本与GetTreeTableIncludeLevelFieldSql()一样
            string comma = " ,";
            string dot = ".";
            StringBuilder sb1 = new StringBuilder(); //定位成员的字段
            sb1.Append(idFiledName).Append(comma)
                .Append(nameFieldName).Append(comma)
                .Append(parentIdName);
            StringBuilder sb2 = new StringBuilder(); //递归成员的字段
            sb2.Append(tblName).Append(dot).Append(idFiledName).Append(comma)
                .Append(tblName).Append(dot).Append(nameFieldName).Append(comma)
                .Append(tblName).Append(dot).Append(parentIdName);

            string char1 = Enumerable.Repeat(rearChat.ToString(), eachSortFieldLength).Aggregate((current, next) => next + current);
            string char2 = Enumerable.Repeat(rearChat.ToString(), reportSortFileTotallength).Aggregate((current, next) => next + current);

            if (otherFiledName != null && otherFiledName.Length > 0)
            {
                foreach (var item in otherFiledName)
                {
                    sb1.Append(item).Append(comma);
                    sb2.Append(tblName).Append(dot).Append(item).Append(comma);
                }
                sb1.RemoveLastChar(comma.Length);
                sb2.RemoveLastChar(comma.Length);
            }

            string sql = $@"with cte as( 
            SELECT {sb1} , 1 as Level , CAST( LEFT(LTRIM({idFiledName})+'{char1}',{eachSortFieldLength}) AS VARCHAR(10)) AS 'Depth'
            FROM {tblName} WHERE {rootItemWhere}
            union all 
            SELECT {sb2} , cte.Level+1 as Level , CAST(LTRIM(cte.Depth) + LEFT(LTRIM({tblName}.{idFiledName}) +'{char1}',{eachSortFieldLength})AS VARCHAR(10)) AS 'Depth' FROM cte, {tblName} 
                where cte.{idFiledName} = {tblName}.{parentIdName} 
            )
            SELECT {sb1} , Level,LEFT(LTRIM(cte.Depth)+'{char2}',{reportSortFileTotallength})  AS 'sort'  FROM cte
            ORDER BY sort ,cte.Level";
            return sql;

        }

        #endregion

    }
}
