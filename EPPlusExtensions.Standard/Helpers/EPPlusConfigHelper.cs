using EPPlusExtensions.ExtensionMethods;
using EPPlusExtensions.Helpers;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace EPPlusExtensions
{
    /// <summary>
    /// 程序集内部方法
    /// </summary>
    internal sealed class EPPlusConfigHelper
    {
        /// <summary>
        /// 填充head
        /// </summary>
        /// <param name="config"></param>
        /// <param name="configSource"></param>
        /// <param name="worksheet"></param>
        public static void FillData_Head(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            if (config.Head.ConfigCellList is null || config.Head.ConfigCellList.Count <= 0)
            {
                return;
            }
            if (configSource.Head.CellsInfoList == null)// 没有数据源
            {
                return;
            }

            var dictConfigSourceHead = configSource.Head.CellsInfoList.ToDictionary(a => a.ConfigValue);

            foreach (var item in config.Head.ConfigCellList)
            {
                if (configSource.Head is null || configSource.Head.CellsInfoList is null ||
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
        /// <exception cref="Exception"></exception>
        public static int FillData_Body(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet)
        {
            //填充body
            var sheetBodyAddRowCount = 0; //新增了几行 (统计sheet body 在原有的模版上新增了多少行), 需要返回的

            if (config is null || configSource is null ||
                config.Body is null || configSource.Body is null ||
                config.Body.ConfigList is null || configSource.Body.ConfigList is null ||
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

                if (datatable is null || datatable.Rows.Count <= 0) //数据源为null或为空
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
                            if (cells.Length != 2)
                            {
                                throw new Exception("该合并单元格的标识有问题,不是类似于A1:A2这个格式的");
                            }
                            int mergeCellStartRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[0]));
                            int mergeCellEndRow = Convert.ToInt32(RegexHelper.GetLastNumber(cells[1]));

                            driftVale = mergeCellEndRow - mergeCellStartRow + 1;
                            if (driftVale <= 0)
                            {
                                throw new Exception("合并单元格的合并行数小于1");
                            }

                            delRow = mergeCellStartRow + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }
                        else //不是合并单元格
                        {
                            delRow = Convert.ToInt32(RegexHelper.GetLastNumber(cellConfigInfo.Address)) + sheetBodyAddRowCount - sheetBodyDeleteRowCount;
                        }

                        if (delRow <= 0)
                        {
                            throw new Exception("要删除的行号不能小于0");
                        }

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
                var hasMergeCell = dictConfig[nth].ConfigLine?.Find(a => a.Address.Contains(":")) != null;
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
                            ExcelWorksheetHelper.SetReport(worksheet, row, config, destRow, maxIntervalRow);
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

                        if (datatable.Rows.Count <= 1)
                        {
                            continue; //1.数据源中的数据行数大于1才增行
                        }

                        if (i <= tempLine - 2)
                        {
                            continue; //i从0开始,这边要-1,然后又要留一行模版,做为复制源,所以这里要-2
                        }

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
                                    int rowFrom;
                                    int rows;
                                    int copyStylesFromRow;
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
                                List<ExcelCellRange> rangeCells = ExcelWorksheetHelper.GetMergedCellFromRow(worksheet, lastSpaceLineRowNumber, leftColStr, rightColStr);
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
                                if (dictConfigSource[nth].FillMethod is null)
                                {
                                    continue;
                                }
                                var fillMethod = dictConfigSource[nth].FillMethod;
                                if (fillMethod is null || fillMethod.FillDataMethodOption == SheetBodyFillDataMethodOption.Default)
                                {
                                    continue;
                                }
                                if (fillMethod.FillDataMethodOption == SheetBodyFillDataMethodOption.SynchronizationDataSource)
                                {
                                    var isFillDataTitle = fillMethod.SynchronizationDataSource.NeedTitle && i == 0;
                                    var isFillDataBody = fillMethod.SynchronizationDataSource.NeedBody;

                                    if (!isFillDataTitle && !isFillDataBody)
                                    {
                                        continue;
                                    }

                                    if (fillDataColumnsStat is null)
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
                                            if (item.State != FillDataColumnsState.WillUse)
                                            {
                                                continue;
                                            }

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
                                            if (item.State != FillDataColumnsState.WillUse)
                                            {
                                                continue;
                                            }

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
                                            //extensionCell_body.Style.HorizontalAlignment = ...
                                            //extensionCell_body.Style.VerticalAlignment = ...
                                            //11.设置整个sheet的背景色
                                            //extensionCell_body.Fill.PatternType = ... //必须设置这个 ExcelFillStyle.Solid;
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
                            ExcelWorksheetHelper.SetReport(worksheet, row, config, destRow);
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
        ///  获得Database数据源的所有的列的使用状态
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="configLine"></param>
        /// <param name="fillModel"></param>
        /// <returns></returns>
        public static Dictionary<string, FillDataColumns> InitFillDataColumnStat(DataTable dataTable, List<EPPlusConfigFixedCell> configLine, SheetBodyFillDataMethod fillModel)
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

        public static void Modify_DataColumnsIsUsedStat(Dictionary<string, FillDataColumns> fillDataColumnsStat, string columns, bool selectColumnIsWillUse)
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
        public static void FillData_Foot(EPPlusConfig config, EPPlusConfigSource configSource, ExcelWorksheet worksheet, int sheetBodyAddRowCount)
        {
            if (config.Foot.ConfigCellList is null || config.Foot.ConfigCellList.Count <= 0)
            {
                return;
            }

            var dictConfigSource = configSource.Foot.CellsInfoList.ToDictionary(a => a.ConfigValue);
            foreach (var item in config.Foot.ConfigCellList)
            {
                if (configSource.Foot is null ||
                    configSource.Foot.CellsInfoList is null ||
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
        /// 设置单元格的的值
        /// </summary>
        /// <param name="config"></param>
        /// <param name="cells">这里用s结尾,表示单元格有可能是合并单元格</param>
        /// <param name="val">值</param>
        /// <param name="colMapperName">excel填充的列名,不想传值请使用null,用来确保填充的数据格式,譬如身份证, 那么单元格必须要是</param>
        public static void SetWorksheetCellsValue(EPPlusConfig config, ExcelRange cells, object val, string colMapperName)
        {
            var cellValue = config.UseFundamentals
                ? config.CellFormatDefault(colMapperName, val, cells)
                : val;

            if (cells.IsRichText)
            {
                cells.RichText.Text = cellValue?.ToString() ?? "";
            }
            else
            {
                cells.Value = cellValue;
            }

            //注:排除3种值( DBNull.Value , null , "") 后 如果 cells.Value 仍然没有值,有可能是配置的单元格以 '开头.
            //譬如: '$tb1Id. 对于这种配置我程序无法检测出来(或者说我没找到检测'开头的方法)
            //下面代码有问题,当遇到日期类型的时候, 值是赋值上去的,但是 cells.value 却!= val .所以下面代码注释
            //if (val != DBNull.Value  && val != null && val != "" && cells.Value != val)
            //{
            //    //如果值没赋值上去,那么抛异常
            //    throw new Exception($"工作簿'{cells.Worksheet.Name}'的配置列'{colMapperName}'的单元格格式有问题,程序无法将值'{val}'赋值到对应的单元 格'{cells.Address}'中.配置的单元格中可能是'开头的,请把'去掉");
            //}
        }

    }
}