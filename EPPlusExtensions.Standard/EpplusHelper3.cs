using EPPlusExtensions.Attributes;
using EPPlusExtensions.Exceptions;
using EPPlusExtensions.Helper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusExtensions
{
    public partial class EPPlusHelper
    {
        #region 对单元格样式进行 Get Set

        ///// <summary>
        /////  获取Cell样式
        ///// </summary>
        ///// <param name="cell"></param>
        ///// <returns></returns>
        //public static EPPlusCellStyle GetCellStyle(ExcelRange cell)
        //{
        //    EPPlusCellStyle cellStyle = new EPPlusCellStyle();
        //    cellStyle.HorizontalAlignment = cell.Style.HorizontalAlignment;
        //    cellStyle.VerticalAlignment = cell.Style.VerticalAlignment;
        //    cellStyle.WrapText = cell.Style.WrapText;
        //    cellStyle.FontBold = cell.Style.Font.Bold;
        //    cellStyle.FontColor = string.IsNullOrEmpty(cell.Style.Font.Color.Rgb)
        //        ? Color.Black
        //        : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Font.Color.Rgb);
        //    cellStyle.FontName = cell.Style.Font.Name;
        //    cellStyle.FontSize = cell.Style.Font.Size;
        //    cellStyle.BackgroundColor = string.IsNullOrEmpty(cell.Style.Fill.BackgroundColor.Rgb)
        //        ? Color.Black
        //        : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Fill.BackgroundColor.Rgb);
        //    cellStyle.ShrinkToFit = cell.Style.ShrinkToFit;
        //    return cellStyle;
        //}

        ///// <summary>
        ///// 设置Cell样式
        ///// </summary>
        ///// <param name="cell"></param>
        ///// <param name="style"></param>
        //public static void SetCellStyle(ExcelRange cell, EPPlusCellStyle style)
        //{
        //    cell.Style.HorizontalAlignment = style.HorizontalAlignment;
        //    cell.Style.VerticalAlignment = style.VerticalAlignment;
        //    cell.Style.WrapText = style.WrapText;
        //    cell.Style.Font.Bold = style.FontBold;
        //    cell.Style.Font.Color.SetColor(style.FontColor);
        //    if (!string.IsNullOrEmpty(style.FontName))
        //    {
        //        cell.Style.Font.Name = style.FontName;
        //    }
        //    cell.Style.Font.Size = style.FontSize;
        //    cell.Style.Fill.PatternType = style.PatternType;
        //    cell.Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
        //    cell.Style.ShrinkToFit = style.ShrinkToFit;
        //}

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
            if (!excelPackage.Compatibility.IsWorksheets1Based)
            {
                workSheetIndex -= 1; //从0开始的, 自己 -1;
            }
            return workSheetIndex;
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
            var mergeCellAddress = ws.MergedCells[row, col];//最准确的合并单元格值
            if (mergeCellAddress == null)
            {
                //不是合并单元格
                return new ExcelCellPoint(row, col).R1C1;
            }
            else
            {
                return mergeCellAddress;
            }
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
            var cell = ws.Cells[row, col];
            if (cell.Merge)
            {
                var mergeCellAddress = ws.MergedCells[row, col];//最准确的合并单元格值
                var mergeCell = new ExcelAddress(mergeCellAddress);
                //var margeCell_Range = new ExcelCellRange(mergeCellAddress);

                var leftCellRow = mergeCell.Start.Row;
                var leftCellCol = mergeCell.Start.Column - 1;
                var leftCellAddress = ws.MergedCells[leftCellRow, leftCellCol];
                if (leftCellAddress == null) //左边的单元格是普通的单元格
                {
                    return new ExcelCellPoint(leftCellRow, leftCellCol).R1C1;
                }
                else
                {
                    ea = new ExcelAddress(leftCellAddress);
                    return ea.Address;
                }
            }
            else
            {
                return new ExcelCellPoint(row, col - 1).R1C1;
            }
        }

        /// <summary>
        /// 获得合并单元格的值 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellText(ExcelWorksheet ws, int row, int col)
        {
            string mergeCellAddress = ws.MergedCells[row, col];
            if (mergeCellAddress == null) return GetCellText(ws, row, col); //不是合并单元格
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
                return cell.Text;//有的单元格通过cell.Text取值会发生异常,但cell.Value却是有值的
            }
            catch (System.NullReferenceException)
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
        /// <param name="filePath"></param>
        /// <param name="fileOutDirectoryName"></param>
        /// <param name="dataConfigInfo"></param>
        /// <param name="cellCustom"></param>
        /// <returns></returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(string filePath, string fileOutDirectoryName, List<ExcelDataConfigInfo> dataConfigInfo, Action<ExcelRange> cellCustom = null)
        {
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
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
        /// <param name="dataConfigInfo"></param>
        /// <param name="cellCustom"></param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, List<ExcelDataConfigInfo> dataConfigInfo, Action<ExcelRange> cellCustom = null)
        {
            ExcelWorksheets wss = excelPackage.Workbook.Worksheets;
            List<DefaultConfig> list = new List<DefaultConfig>();
            var eachCount = 1;
            foreach (var ws in wss)
            {
                int titleLine = 1;
                int titleColumn = 1;
                if (dataConfigInfo != null)
                {
                    var configInfo = dataConfigInfo.Find(a => a.WorkSheetIndex == eachCount);
                    if (configInfo != null)
                    {
                        //titleLine = configInfo.TitleLine;
                        //titleColumn = configInfo.TitleColumn;
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
                    }
                }

                list.Add(FillExcelDefaultConfig(ws, titleLine, titleColumn, cellCustom));
                eachCount++;
            }
            return list;
        }

        /// <summary>
        /// 返回模版的 titleLine 和  titleColumn
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
            object[,] cellsValue = ws.Cells.Value as object[,];
            if (cellsValue == null) throw new ArgumentNullException();
            return GetCellsBy(ws, cellsValue, a => a != null && a.ToString() == value);
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
            using (FileStream fs = new System.IO.FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                return null;
            }
            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("程序报错:");
                if (e.Message != null && e.Message.Length > 0)
                {
                    sb.AppendLine($@"Message:{e.Message}");
                }
                if (e.InnerException != null && e.InnerException.Message != null && e.InnerException.Message.Length > 0)
                {
                    sb.AppendLine($@"InnerExceptionMessage:{e.InnerException.Message}");
                }

                return sb.ToString();
            }
        }

        #endregion
    }
}
