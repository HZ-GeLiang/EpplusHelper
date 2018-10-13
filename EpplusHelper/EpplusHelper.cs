using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style; 


namespace EpplusHelperExtensions
{
    public class EpplusHelper
    {

        //类型参考网址: http://filext.com/faq/office_mime_types.php
        public const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

        /// <summary>
        /// 获得合并单元格的值 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMegerCellText(ExcelWorksheet ws, int row, int col)
        {
            //这个方法的是百度抄来的修改过的 http://blog.csdn.net/xuxiushi888/article/details/49001327
            string range = ws.MergedCells[row, col];
            if (range == null) return GetCellText(ws, row, col);
            var ea = new ExcelAddress(range).Start;
            return ws.Cells[ea.Row, ea.Column].Text;
        }

        #region FillData

        /// <summary>
        /// 往目标sheet中填充数据.注:填充的数据的类型会影响你设置的excel单元格的格式是否起作用
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="workSheetNewName">填充数据后的Worksheet叫什么.  </param>
        /// <param name="destWorkSheetName">填充数据的workSheet叫什么</param> 
        public static void FillData(ExcelPackage excelPackage, EpplusConfig config, EpplusConfigSource configSource, string workSheetNewName, string destWorkSheetName)
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
        public static void FillData(ExcelPackage excelPackage, EpplusConfig config, EpplusConfigSource configSource, string workSheetNewName, int destWorkSheetIndex)
        {

            if (workSheetNewName == null) throw new ArgumentNullException(nameof(workSheetNewName));
            if (destWorkSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(destWorkSheetIndex));
            ExcelWorksheet worksheet = GetExcelWorksheet(excelPackage, destWorkSheetIndex, workSheetNewName);
            config.WorkSheetDefault?.Invoke(worksheet);
            FillData(config, configSource, worksheet);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="cells">s结尾表示单元格有可能是合并单元格</param>
        /// <param name="val">值</param>
        /// <param name="colMapperName">excel填充的列名,不想传值请使用null</param> 
        private static void SetWorksheetCellsValue(EpplusConfig config, ExcelRange cells, object val, string colMapperName)
        {
            if (config.UseFundamentals)
            {
                config.CellFormatDefault(colMapperName, val, cells);
            }
            cells.Value = val;
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
        /// 往目标sheet中填充数据
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        private static void FillData(EpplusConfig config, EpplusConfigSource configSource, ExcelWorksheet worksheet)
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
            //填充body
            int sheetBodyDeleteRowCount = 0; //sheet body 中删除了多少行(只含配置的行,对于FillData()内的删除行则不包括在内).  
            var sheetBodyAddRowCount = 0; //新增了几行 (统计sheet body 在原有的模版上新增了多少行)
            foreach (var nth in config.SheetBodyMapperExcel) //body的第N个配置
            {
                int currentLoopAddLines = 0;
                DataTable datatable;
                if (!configSource.SheetBody.ContainsKey(nth.Key)) //如果没有数据源中没有excle中配置
                {
                    if (config.DeleteFillDateStartLineWhenDataSourceEmpty) //需要删除配置行(当数据源为空[无,null.rows.count=0])
                    {
                        datatable = null;
                    }
                    else
                    {
                        continue; //不许要删除删除配置空行,那么直接跳过
                    }
                }
                else
                {
                    datatable = configSource.SheetBody[nth.Key]; //body的第N个配置的数据源
                }

                if (datatable == null || datatable.Rows.Count <= 0) //数据源为null或为空
                {
                    //throw new ArgumentNullException($"configSource.SheetBody[{nth.Key}]没有可读取的数据");
                    if (config.DeleteFillDateStartLineWhenDataSourceEmpty && nth.Value.Keys.Count > 0)
                    {
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
                            break; //结束删除行循环,因为我只要删一次即可
                        }
                    }
                    continue; //跳过本次fillDate的循环
                }
                int tempLine = 1; //获得第N个配置中excel模版提供了多少行,默认1行
                if (config.SheetBodyMapperExceltemplateLine.ContainsKey(nth.Key))
                {
                    tempLine = config.SheetBodyMapperExceltemplateLine[nth.Key];
                }
                var deleteLastSpaceLine = false; //是否删除最后一空白行(可能有多行组成的)
                int lastSpaceLineInterval = 0; //表示最后一空白行由多少行组成,默认为0
                int lastSpaceLineRowNumber = 0; //表示最后一行的行号是多少

                var hasMergeCell = nth.Value.Keys.ToList().Find(a => a.Contains(":")) != null;
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
                                   ? cellRange[0].Start.Row + (tempLine - 1) * (maxIntervalRow + 1) + sheetBodyAddRowCount - sheetBodyDeleteRowCount
                                   : cellRange[0].Start.Row + i * (maxIntervalRow + 1) + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
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
                                lastSpaceLineRowNumber = destRow + maxIntervalRow + 1;//最后一行空行的位置
                                worksheet.InsertRow(destRow, maxIntervalRow + 1, destRow + maxIntervalRow + 1);//新增N行,注意,此时新增行的高度是有问题的
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
                            string colMapperName = nth.Value[cellRange[j].Range];
                            var val = row[colMapperName];
                            if (cellRange[j].IsMerge)
                            {
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
                            }
                            else
                            {
#if DEBUG
                                throw new Exception("还没遇到这种情况的模版,代码暂时不知道要怎么写.但觉的这种情况应该不会出现");
#else
                return ;
#endif

                            }

                        }
                        if (config.IsReport)
                        {
                            SetReport(worksheet, row, config, destRow, maxIntervalRow);
                        }
                    }
                }
                else //sheet body是常规类型的,即,没有合并单元格的(或者是同行多列的单元格)
                {
                    List<ExcelCellPoint> startCellPointLine = nth.Value.Keys.Select(a => new ExcelCellPoint(a)).ToList(); // 将配置的值 转换成 ExcelCellPoint

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

                        if (datatable.Rows.Count > 1)//1.数据源中的数据行数大于1才增行
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
                                //然后把原本的destRow的样式格式化到新增行中.注意:copy 会把copy行的文本也复制出来.//这里可以说是一个潜在的隐患bug把.因为和我的本意不一样.主要是我不知道要怎么写,只找到一个copy方法,而且copy方法也能帮我解决掉 同一行的 单元格合并问题
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
                        //变量 j 的终止条件不能是 datatable.Rows.Count. 因为datatable可能会包含多余的字段信息,与 配置信息列的个数不一致.
                        for (int j = 0; j < startCellPointLine.Count; j++)
                        {
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

                            bool isAddTitle = false;
                            if (j == startCellPointLine.Count - 1)
                            {
                                if (!configSource.SheetBodySync.ContainsKey(nth.Key))
                                {
                                    continue;
                                }
                                var syncConfig = configSource.SheetBodySync[nth.Key];
                                if (syncConfig == null)
                                {
                                    continue;
                                }
                                if (!syncConfig.SyncSheetBody)
                                {
                                    continue;
                                }

                                var notMapperColumn = new Dictionary<string, bool>();

                                foreach (DataColumn column in datatable.Columns)
                                {
                                    notMapperColumn.Add(column.ColumnName, false);
                                }
                                foreach (var item in nth.Value)
                                {
                                    notMapperColumn[item.Value] = true;
                                }

                                if (!isAddTitle && syncConfig.SyncSheetBodyNeedTitle)
                                {
                                    var extensionDestRowStart_title = nth.Key == 1
                                         ? startCellPointLine[0].Row
                                         : startCellPointLine[0].Row + sheetBodyAddRowCount;

                                    var config_firstCell_col = new ExcelCellPoint(nth.Value.First().Key).Col;
                                    var extensionDestColStart_title = config_firstCell_col + nth.Value.Count;
                                    foreach (var item in notMapperColumn)
                                    {
                                        if (item.Value) continue;

                                        extensionDestColStart_title++;
                                        var extTitleCells = worksheet.Cells[extensionDestRowStart_title, extensionDestColStart_title];
                                        SetWorksheetCellsValue(config, extTitleCells, item.Key, colMapperName);//item.key 是列名

                                    }
                                    isAddTitle = true;


                                }

                                var extensionDestRowStart = nth.Key == 1
                                    ? startCellPointLine[0].Row
                                    : startCellPointLine[0].Row + sheetBodyAddRowCount;

                                var extensionDestColStart = destCol;
                                extensionDestColStart++;

                                var extCells = worksheet.Cells[extensionDestRowStart, extensionDestColStart];
                                SetWorksheetCellsValue(config, cells, val, colMapperName);

                            }
                        }

                        if (config.IsReport)
                        {
                            SetReport(worksheet, row, config, destRow);
                        }
                    }
                }
                //已经修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白.以后有空修复
                if (deleteLastSpaceLine)
                {
                    worksheet.DeleteRow(lastSpaceLineRowNumber, lastSpaceLineInterval, true);
                    sheetBodyAddRowCount -= lastSpaceLineInterval;
                }
                if (config.SheetBodySummaryMapperExcel.Keys.Contains(nth.Key))
                {
                    foreach (var item in config.SheetBodySummaryMapperExcel[nth.Key])//填充第N个配置的一些零散的单元格的值(譬如汇总信息等)
                    {
                        var cellpoint = new ExcelCellPoint(item.Key);
                        // worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col].Value = configSource.SheetBodySummary[nth.Key][item.Value];
                        //item.Key -> A24 , item.Value -> 平均值
                        string colMapperName = item.Value;
                        object val = configSource.SheetBodySummary[nth.Key][item.Value];
                        ExcelRange cells = worksheet.Cells[cellpoint.Row + sheetBodyAddRowCount, cellpoint.Col];

                        if (config.SheetBodySummaryCellCustomSetValue.ContainsKey(nth.Key) && config.SheetBodySummaryCellCustomSetValue[nth.Key] != null)
                        {
                            config.SheetBodySummaryCellCustomSetValue[nth.Key].Invoke(colMapperName, val, cells);
                        }
                        else
                        {

                            SetWorksheetCellsValue(config, cells, val, colMapperName);
                        }
                    }
                }
            }
            //填充foot
            foreach (var item in config.SheetFootMapperExcel)
            {
                if (configSource.SheetFoot.Keys.Count == 0)//excel中有配置foot,但是程序中没有进行值的映射(没映射的原因之一是没有查询出数据)
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

        #endregion

        /// <summary>
        /// 设置报表(能折叠行的excel)格式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="config"></param>
        /// <param name="destRow"></param>
        /// <param name="maxIntervalRow"></param>
        private static void SetReport(ExcelWorksheet worksheet, DataRow row, EpplusConfig config, int destRow, int maxIntervalRow = 0)
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

        public static EpplusConfig GetEmptyConfig()
        {
            return new EpplusConfig();
        }

        public static EpplusConfigSource GetEmptyConfigSource()
        {
            return new EpplusConfigSource();
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
            });
        }

        /// <summary>
        /// 只能是最普通的excel.(每个单元格都是未合并的,第一行是列名,数据从第一列开始填充的那种.)
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
            });
        }


        public static List<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {

            ExcelWorksheet ws = args.ws;
            int rowIndex = args.rowIndex_Data;
            string everyCellPrefix = args.EveryCellPrefix;
            int dataNameRowIndex = args.rowIndex_DataName;
            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplace == null ? args.EveryCellReplaceDefault : args.EveryCellReplace;
            var havingFilter = args.HavingFilter;
            var whereFilter = args.WhereFilter;

            List<T> list = new List<T>();

            var colNameList = GetExcelColumnOfModel(ws, dataNameRowIndex, 1, EpplusConfig.MaxCol07);
            var dictColName = colNameList.ToDictionary(item => item.ExcelCellPoint.Col, item => item.Value.ToString());

            string modelCheckMsg;
            if (!IsAllExcelColumnExistsModel<T>(colNameList, out modelCheckMsg)) throw new ExcelColumnNotExistsWithModelException(modelCheckMsg);

            for (int row = rowIndex; row <= EpplusConfig.MaxRow07; row++)
            {
                if (string.IsNullOrEmpty(ws.Cells[row, 1].Text))//列名为空
                {
                    break;
                }
                Type type = typeof(T);
                var ctor = type.GetConstructor(new Type[] { });
                if (ctor == null) throw new ArgumentException($"通过反射无法得到{type.FullName}的一个无构造参数的构造器:");
                T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转

                for (int col = 1; col < EpplusConfig.MaxCol07; col++)
                {
                    //string colName = ws.Cells[1, col].Text;
                    //去掉不合理的属性命名的字符串(提取合法的字符并接成一个字符串)
                    //string colName = RegexHelper.GetStringByReg(ws.Cells[1, col].Text, @"[_a-zA-Z\u4e00-\u9FFF][A-Za-z0-9_\u4e00-\u9FFF]*").Aggregate("", (current, item) => current + item);
                    if (!dictColName.ContainsKey(col)) break;
                    string colName = dictColName[col];
                    if (string.IsNullOrEmpty(colName)) break;
                    PropertyInfo pInfo = type.GetProperty(colName);

                    string value = ws.Cells[row, col].Text;
                    if (everyCellPrefix?.Length > 0)
                    {
                        var indexof = value.IndexOf(everyCellPrefix);
                        if (indexof == -1)
                        {
                            throw new Exception($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + everyCellPrefix + "'开头的");
                        }
                        value = value.RemovePrefix(everyCellPrefix);
                    }
                    if (everyCellReplace != null)
                    {
                        foreach (var item in everyCellReplace)
                        {
                            var everyCellReplaceOldValue = item.Key;
                            var everyCellReplaceNewValue = item.Value ?? "";
                            if (everyCellReplaceOldValue?.Length > 0)
                            {
                                value = value.Replace(everyCellReplaceOldValue, everyCellReplaceNewValue);
                            }
                        }
                    }

                    if (pInfo.PropertyType == typeof(string))
                    {
                        pInfo.SetValue(model, value);
                        //pInfo.SetValue(model, ws.Cells[row, col].Text);
                        //type.GetProperty(colName)?.SetValue(model, ws.Cells[row, col].Text);
                    }
                    else if (pInfo.PropertyType == typeof(DateTime?))
                    {
                        value = value.Trim();
                        if (value.Length > 0)
                        {
                            pInfo.SetValue(model, Convert.ToDateTime(value));
                        }
                        else
                        {
                            pInfo.SetValue(model, null);
                        }
                    }
                    else if (pInfo.PropertyType == typeof(DateTime))
                    {
                        value = value.Trim();
                        pInfo.SetValue(model, Convert.ToDateTime(value));
                    }
                    else if (pInfo.PropertyType == typeof(decimal?))
                    {
                        value = value.Trim();
                        if (value.Length > 0)
                        {
                            pInfo.SetValue(model, Convert.ToDecimal(value));
                        }
                        else
                        {
                            pInfo.SetValue(model, null);
                        }
                    }
                    else if (pInfo.PropertyType == typeof(decimal))
                    {
                        value = value.Trim();
                        pInfo.SetValue(model, Convert.ToDecimal(value));
                    }
                    else if (pInfo.PropertyType == typeof(Int16))
                    {
                        value = value.Trim();
                        pInfo.SetValue(model, Convert.ToInt16(value));
                    }
                    else if (pInfo.PropertyType == typeof(Int32))
                    {
                        value = value.Trim();
                        pInfo.SetValue(model, Convert.ToInt32(value));
                    }
                    else if (pInfo.PropertyType == typeof(Int64))
                    {
                        value = value.Trim();
                        pInfo.SetValue(model, Convert.ToInt64(value));
                    }
                    else
                    {
                        throw new Exception("未考虑到的情况!!!请完善程序");
                    }
                }

                if (whereFilter == null || whereFilter.Invoke(model))
                {
                    list.Add(model);
                }
            }
            return havingFilter == null ? list : list.Where(item => havingFilter.Invoke(item)).ToList();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">T是给Filter用的</typeparam>
        public class GetExcelListArgs<T> where T : class, new()
        {
            public ExcelWorksheet ws { get; set; }

            /// <summary>
            /// 数据起始行(不含列名),从1开始
            /// </summary>
            public int rowIndex_Data { get; set; } // = 2;

            /// <summary>
            /// 被遍历的单元格内容不为空时的起始字符必须是该字符,然后忽略该字符
            /// </summary>
            public string EveryCellPrefix { get; set; } = "";

            public Dictionary<string, string> EveryCellReplace { get; set; } = null;

            /// <summary>
            /// 数据起始行(不含列名),从1开始
            /// </summary>
            public int rowIndex_DataName { get; set; }

            public bool UseEveryCellReplace { get; set; } = true;

            /// <summary>
            /// EveryCellReplace的默认提供
            /// </summary>
            public Dictionary<string, string> EveryCellReplaceDefault = new Dictionary<string, string>
            {
                {"\t", ""},
                {"\r", ""},
                {"\n", ""},
                {"\r\n", ""},
            };
            /// <summary>
            /// 在return数据之前执行过滤操作
            /// </summary>
            public Func<T, bool> HavingFilter = null;

            /// <summary>
            /// 检查数据,如果数据正确,添加到 返回数据 集合中
            /// </summary>
            public Func<T, bool> WhereFilter = null;

        }


        /// <summary>
        ///  从Excel 中获得符合C# 类属性定义的列名集合
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">列名在Excel的第几行</param>
        /// <returns></returns>
        public static List<CellInfo> GetExcelColumnOfModel(ExcelWorksheet ws, int row, int colStart, int? colEnd)
        {
            if (colEnd == null) colEnd = EpplusConfig.MaxCol07;
            var list = new List<CellInfo>();
            for (int col = colStart; col < colEnd; col++)
            {
                //去掉不合理的属性命名的字符串(提取合法的字符并接成一个字符串)
                string colName = RegexHelper.GetStringByReg(ws.Cells[row, col].Text, @"[_a-zA-Z\u4e00-\u9FFF][A-Za-z0-9_\u4e00-\u9FFF]*")
                    .Aggregate("", (current, item) => current + item);
                if (string.IsNullOrEmpty(colName)) break;

                list.Add(new CellInfo()
                {
                    WorkSheet = ws,
                    Value = colName,
                    ExcelCellPoint = new ExcelCellPoint(row, col)
                });
            }

            return list;
        }

        /// <summary>
        /// excel的所有列均在model对象的属性中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="colNameList"></param>
        /// <param name="modelCheckMsg"></param>
        /// <returns></returns>
        public static bool IsAllExcelColumnExistsModel<T>(List<CellInfo> colNameList, out string modelCheckMsg) where T : class, new()
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

            if (sb.Length > 0)
            {
                modelCheckMsg = sb.ToString();
                return false;
            }
            modelCheckMsg = "";
            return true;
        }


        #region 一些帮助方法

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
        /// 获取excel中对应的值
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static List<CellInfo> GetCellsBy(ExcelWorksheet ws, string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var cellsValue = ws.Cells.Value as object[,];
            if (cellsValue == null) throw new ArgumentNullException();

            return GetCellsBy(ws, cellsValue, a => a != null && a.ToString() == value);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellsValue">一般通过ws.Cells.Value as object[,] 获得 </param>
        /// <param name="match">示例: a => a != null && a.GetType() == typeof(string) && ((string) a == "备注")</param>
        /// <returns></returns>
        public static List<CellInfo> GetCellsBy(ExcelWorksheet ws, object[,] cellsValue, Predicate<object> match)
        {
            if (cellsValue == null) throw new ArgumentNullException(nameof(cellsValue));

            var result = new List<CellInfo>();
            for (int i = 0; i < cellsValue.GetLength(0); i++)
            {
                for (int j = 0; j < cellsValue.GetLength(1); j++)
                {
                    if (match != null && match.Invoke(cellsValue[i, j]))
                    {
                        result.Add(new CellInfo
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


            // return cell.Text; //这个没有科学计数法  注:Text是Excel显示的值,Value是实际值.
        }

        #endregion

        #region 读取excel配置

        /// <summary>
        /// 从Excel中获取默认的配置信息
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetIndex">第几个workSheet作为模版,从1开始</param>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EpplusConfig config, int workSheetIndex)
        {
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
            SetDefaultConfigFromExcel(excelPackage, config, sheet);
        }

        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EpplusConfig config, string workSheetName)
        {
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));
            var sheet = GetExcelWorksheet(excelPackage, workSheetName);
            SetDefaultConfigFromExcel(excelPackage, config, sheet);
        }
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EpplusConfig config, ExcelWorksheet sheet)
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
        public static void SetConfigHeadFromExcel(ExcelPackage excelPackage, EpplusConfig config, ExcelWorksheet sheet)
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
        public static void SetConfigBodyFromExcel(ExcelPackage excelPackage, EpplusConfig config, ExcelWorksheet sheet)
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
                            ? EpplusConfig.MaxRow07 - i
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
                            //后续补充:如果是{"A15:A17", "发生日期"}这种单元格,然后key是A15:K17.在生成excel打开后会提示在
                            //***.xlsx中发现不可读取的内容。是否恢复此工作簿的内容.然后修复后的文档内容是正确的(至少我测试的几个是正确的)
                            //所以,同行多列合并的单元格的key 必须是 A15 这种格式的
                            var newKey = sheetMergedCellsList.Find(a => a.Contains(key));
                            var cells = newKey.Split(':');
                            dictList[nth - 1].Add(
                                RegexHelper.GetFirstNumber(cells[0]) == RegexHelper.GetFirstNumber(cells[1])
                                    ? key
                                    : newKey, val);
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
        public static void SetConfigFootFromExcel(ExcelPackage excelPackage, EpplusConfig config, ExcelWorksheet sheet)
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
                    sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉
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
        public static void SetConfigSourceHead(EpplusConfigSource configSource, DataTable dt, DataRow dr)
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
        public static void SetConfigSourceFoot(EpplusConfigSource configSource, DataTable dt, DataRow dr)
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
            long modifyCellCount = 0;//有几处修改
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

            using (MemoryStream ms = new MemoryStream())
            {
                excelPackage.SaveAs(ms); // 导入数据到流中 
                ms.Position = 0;
                //if (File.Exists(savePath))
                //{
                //    File.Delete(savePath);
                //}
                File.Delete(savePath); //删除文件。如果文件不存在,也不报错
                ms.Save(savePath);
            }

            return modifyCellCount > 0;
        }

        /// <summary>
        ///  科学计数法的cell转成文本格式的cell
        /// </summary>
        /// <param name="fileFullPath">文件路径</param>
        /// <param name="fileSaveAsPath">文件另存为路径</param>
        /// <param name="containNoMatchCell">包含不匹配的单元格(即把所有的单元格变成文本格式),true:是.false:仅针对显示成科学计数法的cell变成文本</param>
        /// <returns>是否有进行科学技术法的cell转换.true:是,fale:否</returns>
        public static bool ScientificNotationFormatToString(string fileFullPath, string fileSaveAsPath, bool containNoMatchCell = false)
        {
            using (FileStream fs = File.OpenRead(fileFullPath))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                return ScientificNotationFormatToString(excelPackage, fileSaveAsPath, containNoMatchCell);
            }
        }

        #endregion

        /// <summary>
        /// 设置Cell样式(没用过,方法体中的代码是大熊发我的)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        public static void SetCellStyle(ExcelRange cell, EpplusCellStyle style)
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

        /// <summary>
        ///  获取Cell样式(没用过,方法体中的代码是大熊发我的)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static EpplusCellStyle GetCellStyle(ExcelRange cell)
        {
            EpplusCellStyle cellStyle = new EpplusCellStyle();
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

        /// <summary>
        /// 我自定义的Cell样式类(根据上面2个方法推出来的)
        /// </summary>
        public class EpplusCellStyle
        {
            public ExcelHorizontalAlignment HorizontalAlignment { get; set; }
            public ExcelVerticalAlignment VerticalAlignment { get; set; }
            public bool WrapText { get; set; }
            public bool FontBold { get; set; }
            public Color FontColor { get; set; }
            public Color BackgroundColor { get; set; }
            public string FontName { get; set; }
            public float FontSize { get; set; }
            public ExcelFillStyle PatternType { get; set; }
            public bool ShrinkToFit { get; set; }
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
                if (EpplusHelper.GetCellText(ws, cell.Row, cell.Col) != dict[key])
                {
                    return false;
                }
            }
            return true;
        }

    }

    /// <summary>
    /// 保存一个合并单元格的结构体
    /// </summary>
    public struct ExcelCellRange
    {
        public ExcelCellRange(string range)
        {
            Range = range;
            var cellPoints = range.Split(':');

            switch (cellPoints.Length)
            {
                case 1:
                    Start = new ExcelCellPoint(cellPoints[0].Trim());
                    End = default(ExcelCellPoint);
                    IntervalCol = 0;
                    IntervalRow = 0;
                    IsMerge = false;
                    break;
                case 2:
                    Start = new ExcelCellPoint(cellPoints[0].Trim());
                    End = new ExcelCellPoint(cellPoints[1].Trim());
                    IntervalCol = End.Col - Start.Col;
                    IntervalRow = End.Row - Start.Row;
                    IsMerge = true;
                    break;
                default:
                    throw new Exception("程序的配置有问题");
            }

        }

        /// <summary>
        /// 范围(保存的是配置时的字符串.在程序中用来当作key使用)
        /// </summary>
        public string Range { get; private set; }

        /// <summary>
        /// 开始Point
        /// </summary>
        public ExcelCellPoint Start { get; private set; }

        /// <summary>
        /// 结束Point
        /// </summary>
        public ExcelCellPoint End { get; private set; }

        /// <summary>
        /// 间距行是多少
        /// </summary>
        public int IntervalRow { get; private set; }

        /// <summary>
        /// 间距列是多少
        /// </summary>
        public int IntervalCol { get; private set; }

        /// <summary>
        /// 是否是合并单元格
        /// </summary>
        public bool IsMerge { get; private set; }
    }

    /// <summary>
    /// 记录ExcelCell位置的一个结构体
    /// </summary>
    public struct ExcelCellPoint
    {
        public int Row;
        public int Col;

        /// <summary>
        /// 譬如A2等
        /// </summary>
        public string R1C1;

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="row">从1开始的整数</param>
        ///// <param name="col">只能是字母</param>
        ///// <param name="r1C1">譬如A2 等</param>
        //public ExcelCellPoint(int row, string col, string r1C1)
        //{
        //    Row = row;
        //    Col = R1C1Formulas(col);
        //    R1C1 = r1C1;
        //}
        public ExcelCellPoint(string r1C1)
        {
            //K3 = row:3, col:11
            r1C1 = r1C1.Split(':')[0].Trim(); //防止传入 "A1:B3" 这种的配置格式的
            Row = Convert.ToInt32(RegexHelper.GetLastNumber(r1C1));//3
            Col = R1C1Formulas(RegexHelper.GetFirstStringByReg(r1C1, "[A-Za-z]+"));//K -> 11
            R1C1 = r1C1;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row">从1开始的整数</param>
        /// <param name="col">从1开始的整数</param>
        public ExcelCellPoint(int row, int col)
        {
            Row = row;
            Col = col;
            R1C1 = R1C1FormulasReverse(col) + row;
        }

        /// <summary>
        /// 譬如: A->1 . 在excel的选项->属性->公式  下有个 R1C1引用样式
        /// </summary>
        /// <param name="col">只能是字母</param>
        /// <returns></returns>
        public static int R1C1Formulas(string col)
        {
            col = col.ToUpper();
            Dictionary<string, int> r1C1 = new Dictionary<string, int>
            {
                {"A", 1},{"B", 2},{"C", 3},{"D", 4},{"E", 5},{"F", 6},
                {"G", 7},{"H", 8},{"I", 9},{"J", 10},{"K", 11},{"L", 12},
                {"M", 13},{"N", 14},{"O", 15},{"P", 16},{"Q", 17},{"R", 18},
                {"S", 19},{"T", 20},{"U", 21},{"V", 22},{"W", 23},{"X", 24},
                {"Y", 25},{"Z", 26},
            };
            int colLength = col.Length;
            if (colLength == 1)
            {
                return r1C1[col];
            }
            int sum = 0;
            for (int i = 0; i < colLength; i++)
            {
                char c = col[i];
                int num = r1C1[c + ""];
                sum += (int)(num * Math.Pow(26, colLength - i - 1));
            }
            return sum;
        }

        /// <summary>
        /// 譬如1->A 
        /// </summary>
        /// <param name="num">excel的第几列</param>
        /// <returns></returns>
        public static string R1C1FormulasReverse(int num)
        {
            if (num <= 0)
            {
                throw new Exception("parameter 'col' can not less zero");
            }
            Dictionary<int, char> r1C1 = new Dictionary<int, char>
            {
                {1,'A'},{2,'B'},{3,'C'},{4,'D'},{5,'E'},{6,'F'},
                {7,'G'},{8,'H'},{9,'I'},{10,'J'},{11,'K'},{12,'L'},
                {13,'M'},{14,'N'},{15,'O'},{16,'P'},{17,'Q'},{18,'R'},
                {19,'S'},{20,'T'},{21,'U'},{22,'V'},{23,'W'},{24,'X'},
                {25,'Y'},{26,'Z'},{0,'A'}
            };
            if (num <= 26) //这个if属于优化,若删掉.也没有关系
            {
                return r1C1[num].ToString();
            }
            List<char> charList = new List<char>();
            int cimi = -1; //次幂数

            while (true)
            {
                cimi++;
                var num2 = (long)Math.Pow(26, cimi); //while的终止条件与计算条件
                if (num >= num2)
                {
                    int mod = num / (int)num2 % 26;//余数
                    charList.Add(mod != 0 ? r1C1[mod] : r1C1[26]);
                    num -= (int)num2;
                }
                else
                {
                    break;
                }
            }

            StringBuilder sb = new StringBuilder();
            for (int i = charList.Count - 1; i >= 0; i--)
            {
                sb.Append(charList[i]);
            }
            return sb.ToString();

        }
    }

    public class CellInfo
    {
        public ExcelWorksheet WorkSheet { get; set; }
        public ExcelCellPoint ExcelCellPoint { get; set; }
        public object Value { get; set; }
    }
    /// <summary>
    /// 配置信息 
    /// </summary>
    public class EpplusConfig
    {
        #region Excel的最大行与列
        /// <summary>
        /// Excel2007开始最大行是1048576,2^20次方
        /// </summary>
        public static readonly int MaxRow07 = 1048576;

        /// <summary>
        /// excel 2007 和excel 2010最大有2^20=1048576行,2^14=16384列
        /// </summary>
        public static readonly int MaxCol07 = 16384;

        /// <summary>
        /// Excel2003的最大行是65536行
        /// </summary>
        public static readonly int MaxRow03 = 65536;

        /// <summary>
        ///  excel 2003 工作表最大有2^16=65536行,2^8=256列
        /// </summary>
        public static readonly int MaxCol03 = 256;
        #endregion

        /// <summary>
        /// 用来初始化的一些数据的
        /// </summary>
        public EpplusConfig()
        {
            SheetHeadMapperExcel = new Dictionary<string, string>();
            SheetHeadMapperSource = new Dictionary<string, string>();
            SheetHeadCellCustomSetValue = null;

            //注:body是没有数据源的配置的,全靠一个默认约定
            SheetBodyMapperExcel = new Dictionary<int, Dictionary<string, string>>();
            SheetBodyCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
            SheetBodySummaryMapperExcel = new Dictionary<int, Dictionary<string, string>>();
            SheetBodySummaryMapperSource = new Dictionary<int, Dictionary<string, string>>();
            SheetBodySummaryCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
            SheetBodyMapperExceltemplateLine = new Dictionary<int, int>();

            SheetFootMapperExcel = new Dictionary<string, string>();
            SheetFootMapperSource = new Dictionary<string, string>();
            SheetFootCellCustomSetValue = null;

            Report = new EpplusReport();
            IsReport = false;
            DeleteFillDateStartLineWhenDataSourceEmpty = false;
        }

        #region head
        /// <summary>
        /// sheet head 用来完成指定单元格的内容配置
        /// 譬如A2,Name. key不区分大小写,即A2与a2是一样的.建议大写
        /// </summary>
        public Dictionary<string, string> SheetHeadMapperExcel { get; set; }

        /// <summary>
        /// sheet head 的数据源的配置
        /// 譬如Name,张三. key严格区分大小写
        /// </summary>
        public Dictionary<string, string> SheetHeadMapperSource { get; set; }
        /// <summary>
        /// 自定义设置值
        /// </summary>
        public Action<string, object, ExcelRange> SheetHeadCellCustomSetValue;

        #endregion

        #region body
        /// <summary>
        /// sheet body 的内容配置.注.int必须是从1开始的且递增+1的自然数
        /// </summary>
        public Dictionary<int, Dictionary<string, string>> SheetBodyMapperExcel;
        /// <summary>
        /// 自定义设置值
        /// </summary>
        public Dictionary<int, Action<string, object, ExcelRange>> SheetBodyCellCustomSetValue;

        /// <summary>
        /// sheet body中固定的单元格. 譬如汇总信息等.譬如A8,Name,前面的int表示这个汇总是哪个SheetBody的
        /// </summary>
        public Dictionary<int, Dictionary<string, string>> SheetBodySummaryMapperExcel { get; set; }

        /// <summary>
        /// sheet body中固定的单元格的数据源,譬如Name,张三
        /// </summary>
        public Dictionary<int, Dictionary<string, string>> SheetBodySummaryMapperSource { get; set; }
        /// <summary>
        /// 自定义设置值
        /// </summary>
        public Dictionary<int, Action<string, object, ExcelRange>> SheetBodySummaryCellCustomSetValue;

        /// <summary>
        /// SheetBody模版自带(提供)多少行(根据这个,在结合数据源,程序内部判断是否新增行)
        /// </summary>
        public Dictionary<int, int> SheetBodyMapperExceltemplateLine { get; set; }
        #endregion

        #region foot

        /// <summary>
        /// sheet foot 用来完成指定单元格的内容配置
        /// 譬如A8,Name
        /// </summary>
        public Dictionary<string, string> SheetFootMapperExcel { get; set; }

        /// <summary>
        /// sheet foot 的数据源
        /// 譬如Name,张三</summary>
        public Dictionary<string, string> SheetFootMapperSource { get; set; }
        /// <summary>
        /// 自定义设置值
        /// </summary>
        public Action<string, object, ExcelRange> SheetFootCellCustomSetValue;
        #endregion

        /// <summary>
        /// 报表(excel能折叠的那种)的显示的一些配置
        /// </summary>
        public EpplusReport Report { get; set; }

        /// <summary>
        /// 标识是否是一个报表格式(excel能折叠的)的Worksheet(目前该属性表示每一个worksheet), 默认False
        /// </summary>
        public bool IsReport { get; set; }

        /// <summary>
        /// 当填充的数据源为空时,是否删除填充的起始行,默认false
        /// </summary>
        public bool DeleteFillDateStartLineWhenDataSourceEmpty;

        /// <summary>
        /// 是否使用默认(单元格格式)约定,默认true 注:settingCellFormat若与默认的发成冲突,会把默认的cell格式给覆盖.
        /// </summary>
        public bool UseFundamentals = true;
        /// <summary>
        /// 默认的单元格格式设置,colMapperName 是配置单元格的名字 譬如 $tb1Id, 那么colMapperName值就为Id
        /// </summary>

        public Action<string, object, ExcelRange> CellFormatDefault = (colMapperName, val, cells) =>
        {
            //关于格式,可以参考右键单元格->设置单元格格式->自定义中的类型 或看文档: https://support.office.microsoft.com/zh-CN/excel ->自定义 Excel->创建或删除自定义数字格式
            string formatStr = cells.Style.Numberformat.Format;
            //含有Id的列,默认是文本类型,目的是放防止出现科学计数法
            if (colMapperName != null && colMapperName.ToLower().EndsWith("id"))
            {
                if (formatStr != "@")
                {
                    cells.Style.Numberformat.Format = "@"; //Format as text
                }
                val = val.ToString(); //确保值是string类型的
            }
            //若没有设置日期格式,默认是yyyy-mm-dd
            //大写字母是为了冗错.注:excel的日期格式写成大写的是不会报错的,但文档中全是小写的.
            var dateCode = new List<char> { '@', 'y', 'Y', 's', 'S', 'm', 'M', 'h', 'H', 'd', 'D', 'A', 'P', ':', '.', '0', '[', ']' };
            if (val is DateTime)
            {
                bool changeformat = true;
                foreach (var c in formatStr) //这边不能用优化成linq,优化成linq有问题
                {
                    if (dateCode.Contains(c))
                    {
                        changeformat = false;
                        break;
                    }
                }
                if (changeformat) //若为true,表示没有人为的设置该cell的日期显示格式
                {
                    cells.Style.Numberformat.Format = "yyyy-mm-dd"; //默认显示的格式
                }
            }
        };

        public Action<ExcelWorksheet> WorkSheetDefault;
        //= worksheet =>
        //{
        //    //worksheet.DefaultColWidth = 72; //默认列宽
        //    //worksheet.DefaultRowHeight = 18; //默认行高
        //    //worksheet.TabColor = Color.Blue; //Sheet Tab的颜色
        //    //worksheet.Cells.Style.WrapText = true; //单元格文字自动换行
        //};



    }

    public class EpplusReport
    {
        /// <summary>
        /// 行折叠的deth(Level)在DataRow中什么列表示.默认:Level
        /// </summary>
        public string RowLevelColumnName { get; set; } = "Level";

        #region 行折叠默认配置

        /// <summary>
        /// 行是否合并显示
        /// </summary>
        public bool Collapsed { get; set; } = true;

        /// <summary>
        /// 合并/展开 行 的折叠符号是否在右边
        /// </summary>
        public bool OutLineSummaryBelow { get; set; } = false;

        #endregion

        #region 列 折叠默认配置(暂不支持该功能,主要是还没遇到列合并的导出需求)

        ///// <summary>
        ///// 列 是否合并显示
        ///// </summary>
        //public bool RowCollaspsed { get; set; } = true;
        ///// <summary>
        ///// 合并/展开 列 的折叠符号是否在下面
        ///// </summary>
        //public bool OutLineSummaryRight { get; set; } = false;

        #endregion
    }

    public class SheetBodySync
    {
        /// <summary>
        /// 与数据源进行同步,当数据源的数据列多余excel的配置时生效
        /// </summary>
        public bool SyncSheetBody { get; set; } = false;

        public bool SyncSheetBodyNeedTitle { get; set; } = true;
    }
    /// <summary>
    /// 配置信息的数据源
    /// </summary>
    public class EpplusConfigSource
    {
        public EpplusConfigSource()
        {
            SheetHead = new Dictionary<string, string>();
            SheetBody = new Dictionary<int, DataTable>();
            SheetBodySync = new Dictionary<int, SheetBodySync>();
            SheetBodySummary = new Dictionary<int, Dictionary<object, object>>();
            SheetFoot = new Dictionary<string, string>();
        }
        public Dictionary<string, string> SheetHead { get; set; }
        public Dictionary<int, DataTable> SheetBody { get; set; }
        public Dictionary<int, SheetBodySync> SheetBodySync { get; set; }
        public Dictionary<int, Dictionary<object, object>> SheetBodySummary { get; set; }
        public Dictionary<string, string> SheetFoot { get; set; }

    }

    /// <summary>
    /// excel列不在Model中异常
    /// </summary>
    public class ExcelColumnNotExistsWithModelException : Exception
    {
        public ExcelColumnNotExistsWithModelException() : base() { }
        public ExcelColumnNotExistsWithModelException(string message) : base(message) { }
    }
}
