using EPPlusExtensions.ExtensionMethods;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace EPPlusExtensions
{
    public sealed class ExcelWorksheetHelper
    {
        /// <summary>
        /// 获得精确的合并单元格地址
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, int row, int col)
        {
            return ExcelWorksheetHelper.IsMergeCell(ws, row, col, out var mergeCellAddress)
                ? mergeCellAddress
                : new ExcelCellPoint(row, col).R1C1;
        }

        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, string r1c1)
        {
            var excelRange = new ExcelCellRange(r1c1);
            if (excelRange.End.Col == 0) //r1c1 为 D4  这种值
            {
                return ExcelWorksheetHelper.GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.Start.Col);
            }
            else
            {
                return ExcelWorksheetHelper.GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.End.Col);
            }
        }

        public static string GetLeftCellAddress(ExcelWorksheet ws, string currentCellAddress)
        {
            var ea = new ExcelAddress(currentCellAddress);
            var row = ea.Start.Row;
            var col = ea.Start.Column;

            if (!ExcelWorksheetHelper.IsMergeCell(ws, row: row, col: col, out var mergeCellAddress))
            {
                return new ExcelCellPoint(row, col - 1).R1C1;
            }

            var mergeCell = new ExcelAddress(mergeCellAddress);
            var leftCellRow = mergeCell.Start.Row;
            var leftCellCol = mergeCell.Start.Column - 1;
            if (ExcelWorksheetHelper.IsMergeCell(ws, row: leftCellRow, col: leftCellCol, out var leftCellAddress))
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
            return ExcelWorksheetHelper.IsMergeCell(ws, row, col, out var mergeCellAddress);
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
            var isMergeCell = ExcelWorksheetHelper.IsMergeCell(ws, row, col, out var mergeCellAddress);
            if (isMergeCell == false)
            {
                return GetCellText(ws, row, col);
            }
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
            return ExcelRangeHelper.GetCellText(ws.Cells[row, col], when_TextProperty_NullReferenceException_Use_ValueProperty);
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
                if (ExcelWorksheetHelper.GetCellText(ws, cell.Row, cell.Col) != dict[key])
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
        /// <param name="ws"></param>
        /// <param name="rowStartIndex">从1开始</param>
        /// <param name="rowEndIndex">最大值:EPPlusConfig.MaxRow07</param>
        /// <param name="action">一般用于修改Hidden状态</param>
        /// <returns></returns>
        public static void EachHiddenRow(ExcelWorksheet ws, int rowStartIndex, int rowEndIndex, Action<ExcelRow> action)
        {
            if (action is null)
            {
                return;
            }

            if (rowEndIndex > EPPlusConfig.MaxRow07)
            {
                rowEndIndex = EPPlusConfig.MaxRow07;
            }
            for (int rowStart = rowStartIndex; rowStart <= rowEndIndex; rowStart++)
            {
                if (ws.Row(rowStart).Hidden)
                {
                    action.Invoke(ws.Row(rowStart));
                }
            }
        }

        public static bool HaveHiddenRow(ExcelWorksheet ws, int rowStartIndex = 1, int rowEndIndex = EPPlusConfig.MaxRow07)
        {
            if (rowEndIndex > EPPlusConfig.MaxRow07)
            {
                rowEndIndex = EPPlusConfig.MaxRow07;
            }
            for (int rowIndex = rowStartIndex; rowIndex <= rowEndIndex; rowIndex++)
            {
                if (ws.Row(rowIndex).Hidden)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="columnStartIndex">从1开始</param>
        /// <param name="columnEndIndex">最大值:EPPlusConfig.MaxCol07</param>
        /// <param name="action">一般用于修改Hidden状态</param>
        /// <returns></returns>
        public static void EachHiddenColumn(ExcelWorksheet ws, int columnStartIndex, int columnEndIndex, Action<ExcelRow> action)
        {
            if (action is null)
            {
                return;
            }

            if (columnEndIndex > EPPlusConfig.MaxCol07)
            {
                columnEndIndex = EPPlusConfig.MaxCol07;
            }
            for (int columnIndex = columnStartIndex; columnIndex <= columnEndIndex; columnIndex++)
            {
                if (ws.Column(columnIndex).Hidden)
                {
                    action.Invoke(ws.Row(columnIndex));
                }
            }
        }

        public static bool HaveHiddenColumn(ExcelWorksheet ws, int columnStartIndex = 1, int columnEndIndex = EPPlusConfig.MaxCol07)
        {
            if (columnEndIndex > EPPlusConfig.MaxCol07)
            {
                columnEndIndex = EPPlusConfig.MaxCol07;
            }
            for (int columnIndex = columnStartIndex; columnIndex <= columnEndIndex; columnIndex++)
            {
                if (ws.Column(columnIndex).Hidden)
                {
                    return true;
                }
            }

            return false;
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
            if (value is null)
            {
                throw new ArgumentNullException(nameof(value));
            }
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
            if (cellsValue is null)
            {
                throw new ArgumentNullException(nameof(cellsValue));
            }

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


        /// <summary>
        /// 所有的合并单元格
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="lineNo">行号</param>
        /// <param name="leftCol">最左边的</param>
        /// <param name="rightCol">最右边的,如果最右边的合并单元格,取合并单元格的最右边列的地址</param>
        /// <returns></returns>
        public static List<ExcelCellRange> GetMergedCellFromRow(ExcelWorksheet worksheet, int lineNo, string leftCol, string rightCol)
        {
            var allCell = ExcelWorksheetHelper.GetCellFromRow(worksheet, lineNo, leftCol, rightCol);

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
        public static List<object> GetCellFromRow(ExcelWorksheet worksheet, int lineNo, string leftCol, string rightCol)
        {
            var leftAddressCol = ExcelCellPoint.R1C1Formulas(leftCol);
            var rightAddressCol = ExcelCellPoint.R1C1Formulas(rightCol);

            var allCell = new List<object>();
            while (true)
            {
                if (ExcelWorksheetHelper.IsMergeCell(worksheet, row: lineNo, col: leftAddressCol, out var mergeCellAddress))
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
        /// 设置报表(能折叠行的excel)格式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="config"></param>
        /// <param name="destRow"></param>
        /// <param name="maxIntervalRow"></param>
        internal static void SetReport(ExcelWorksheet worksheet, DataRow row, EPPlusConfig config, int destRow, int maxIntervalRow = 0)
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
    }
}