using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EPPlusExtensions;
using EPPlusTool;
using EPPlusTool.MethodExtension;

namespace EPPlusHelperTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void TextBoxDragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            ((System.Windows.Forms.TextBox)sender).Text = path;
            LoadDgv(sender, e);
        }

        private void TextBoxDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Link : DragDropEffects.None;
        }

        private static ExcelWorksheet GetWorkSheet(ExcelPackage excelPackage, string wsIndexString)
        {
            if (excelPackage.Workbook.Worksheets.Count == 1)
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
            }
            if (Int32.TryParse(wsIndexString, out int wsIndexInt))
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, wsIndexInt);
            }
            if (EPPlusHelper.GetExcelWorksheetNames(excelPackage).Contains(wsIndexString))
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, wsIndexString);
            }

            throw new ArgumentException("无法打开Excel的Worksheet");
        }

        private void GenerateConfiguration_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }

                var fileDir = Path.GetDirectoryName(filePath);

                var dataConfigInfo = new List<ExcelDataConfigInfo>() { GetExcelDataConfigInfo() };

                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileDir, dataConfigInfo, cell =>
                {
                    var cellValue = EPPlusHelper.GetCellText(cell);

                    foreach (var key in EPPlusHelper.KeysTypeOfDateTime.Where(item => cellValue.Contains(item)))
                    {
                        cell.Style.Numberformat.Format = "yyyy-mm-dd"; //默认显示的格式
                        break;
                    }

                    foreach (var key in EPPlusHelper.KeysTypeOfString.Where(item => cellValue.Contains(item)))
                    {
                        cell.Style.Numberformat.Format = "@"; //Format as text
                        break;
                    }

                    foreach (var key in EPPlusHelper.KeysTypeOfDecimal.Where(item => cellValue.Contains(item)))
                    {
                        //cell.Style.Numberformat.Format = "@"; //Format as text
                        break;
                    }
                });

                var haveConfig = defaultConfigList.Find(a => a.ClassPropertyList.Count > 0) != null;
                if (!haveConfig)
                {
                    MessageBox.Show("未检测到配置信息");
                    return;
                }
                MessageBox.Show($"文件已经生成,在目录'{fileDir}'");
                //if (!fileDir.Contains($@"\Desktop\"))
                //{
                //    WinFormHelper.OpenDirectory(fileDir);
                //}
            });
        }

        private void GenerateConfigurationCode_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }
                var dataConfigInfo = new List<ExcelDataConfigInfo>();
                for (int i = 0; i < dgv1.Rows.Count; i++)
                {
                    dataConfigInfo.Add(new ExcelDataConfigInfo()
                    {
                        WorkSheetIndex = i + 1,
                        TitleLine = Convert.ToInt32(dgv1.Rows[i].Cells[2].Value),
                        TitleColumn = Convert.ToInt32(dgv1.Rows[i].Cells[3].Value)
                    });
                }

                string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));
                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileOutDirectoryName, dataConfigInfo);

                var haveConfig = defaultConfigList.Find(a => a.ClassPropertyList.Count > 0) != null;
                if (!haveConfig)
                {
                    MessageBox.Show("未检测到配置信息");
                    return;
                }

                //将字符串写入文件
                StringBuilder errMsg = new StringBuilder();
                var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result";

                foreach (var item in defaultConfigList)
                {
                    if (item.ClassPropertyList.Count > 0)
                    {
                        File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDataTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDataTableSnippe);
                        File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                    }
                    else
                    {
                        errMsg.Append(item.WorkSheetName + "、");
                    }
                }

                errMsg.RemoveLastChar('、');
                if (errMsg.Length > 0)
                {
                    MessageBox.Show($@"下列工作簿未生成配置项:{errMsg}");
                }
                if (!filePath.GetDirectoryName().Contains(@"\Desktop\"))
                {
                    WinFormHelper.OpenFilePath(filePath.GetDirectoryName());
                }
            });
        }

        private void CheckTemplateConfiguration_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                var ws1Path = this.filePath1.Text.Trim().移除路径前后引号();
                var ws2Path = this.filePath2.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(ws1Path))
                {
                    MessageBox.Show("路径1不能为空");
                    return;
                }
                if (string.IsNullOrEmpty(ws2Path))
                {
                    MessageBox.Show("路径2不能为空");
                    return;
                }
                if (ws1Path == ws2Path)
                {
                    MessageBox.Show("比较文件路径一致,无法比较");
                    return;
                }

                var ws1Index_string = this.wsNameOrIndex1.Text.Trim();
                var ws2Index_string = this.wsNameOrIndex2.Text.Trim();

                var ws1TitleLine = Convert.ToInt32(this.TitleLine1.Text.Trim());
                var ws2TitleLine = Convert.ToInt32(this.TitleLine2.Text.Trim());

                var ws1TitleCol = Convert.ToInt32(this.TitleCol1.Text.Trim());
                var ws2TitleCol = Convert.ToInt32(this.TitleCol2.Text.Trim());

                using (var fs1 = new FileStream(ws1Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var fs2 = new FileStream(ws2Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage1 = new ExcelPackage(fs1))
                using (var excelPackage2 = new ExcelPackage(fs2))
                {
                    var ws1 = GetWorkSheet(excelPackage1, ws1Index_string);
                    var ws2 = GetWorkSheet(excelPackage2, ws2Index_string);
                    var ws1Props = EPPlusHelper.FillExcelDefaultConfig(ws1, ws1TitleLine, ws1TitleCol).ClassPropertyList;
                    var ws2Props = EPPlusHelper.FillExcelDefaultConfig(ws2, ws2TitleLine, ws2TitleCol).ClassPropertyList;
                    {
                        StringBuilder sb = new StringBuilder();
                        AppendCols(ws1Props, ws2Props, sb);
                        if (sb.Length > 1)
                        {
                            MessageBox.Show($@"A与B比较:B未提供列:{sb.RemoveLastChar()}");
                            return;
                        }
                    }

                    {
                        StringBuilder sb = new StringBuilder();
                        AppendCols(ws2Props, ws1Props, sb);

                        if (sb.Length > 1)
                        {
                            MessageBox.Show($@"A与B比较:B多提供列:{sb.RemoveLastChar()}");
                            return;
                        }
                    }

                    MessageBox.Show("A与B比较:内容一致");
                }
            });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="src">源</param>
        /// <param name="compareObj">比较对象</param>
        /// <param name="sb1">输出的内容</param>
        private static void AppendCols(List<ExcelCellInfoValue> src, List<ExcelCellInfoValue> compareObj, StringBuilder sb1)
        {
            foreach (var item in src)
            {
                if (!item.IsRename)
                {
                    if (compareObj.Find(a => a.Name == item.Name) == null)
                    {
                        sb1.Append($@"{item.Name},");
                    }
                }
                else
                {
                    if (compareObj.Find(a => a.Name == item.Name && a.ExcelColNameIndex == item.ExcelColNameIndex) == null)
                    {
                        sb1.Append($@"{item.Name},");
                    }
                }
            }

        }

        private void Btn_SelectExcelFile(object sender, EventArgs e)
        {
            var selectFilePath = WinFormHelper.SelectFile("excel (*.xlsx)|*.xlsx");
            if (selectFilePath.Length > 0)
            {
                if (((System.Windows.Forms.Control)sender).Name == "SelectFileBtn1")
                {
                    this.filePath1.Text = selectFilePath;
                }
                if (((System.Windows.Forms.Control)sender).Name == "SelectFileBtn2")
                {
                    this.filePath2.Text = selectFilePath;
                }
            }
        }


        private void LoadDgv(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                var callerName = ((System.Windows.Forms.Control)sender).Name;
                string filePath = "";
                //if (sender is System.Windows.Forms.TextBox)
                //{
                //    filePath = ((System.Windows.Forms.TextBox)sender).Text.Trim().移除路径前后引号();
                //}
                if (callerName == "filePath1" || callerName == "BtnAnalyze1")
                {
                    filePath = this.filePath1.Text.Trim().移除路径前后引号();
                }
                else if (callerName == "filePath2" || callerName == "BtnAnalyze2")
                {
                    filePath = this.filePath2.Text.Trim().移除路径前后引号();
                }

                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }
                if (string.Compare(".xlsx", System.IO.Path.GetExtension(filePath), true) != 0)
                {
                    MessageBox.Show("只支持.xlsx文件");
                    return;
                }
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    SetDataSourceForDGV(excelPackage, callerName, this);
                    if ((callerName == "filePath1" || callerName == "BtnAnalyze1"))
                    {
                        if (EPPlusHelper.GetWorkSheetNames(excelPackage, eWorkSheetHidden.Hidden, eWorkSheetHidden.VeryHidden).Count > 0)
                        {
                            MessageBox.Show("检测到当前Excel含有隐藏工作簿,建议删除所有隐藏工作簿");
                        }

                        foreach (var ws in EPPlusHelper.GetExcelWorksheets(excelPackage))
                        {
                            var haveHiddenRow = EPPlusHelper.HaveHiddenRow(ws);
                            var haveHiddenColumn = EPPlusHelper.HaveHiddenColumn(ws);

                            if (haveHiddenRow && haveHiddenColumn)
                            {
                                MessageBox.Show($"检测到Sheet页'{ws.Name}'含有隐藏行和隐藏列");
                            }
                            else
                            {
                                if (haveHiddenRow)
                                {
                                    MessageBox.Show($"检测到Sheet页'{ws.Name}'含有隐藏行");
                                }
                                if (haveHiddenColumn)
                                {
                                    MessageBox.Show($"检测到Sheet页'{ws.Name}'含有隐藏列");
                                }
                            }
                        }
                    }

                }
            });
        }

        private static void SetDataSourceForDGV(ExcelPackage excelPackage, string callerName, Form1 form1)
        {
            DataGridView control;
            int block;

            if (callerName == "filePath1" || callerName == "BtnAnalyze1")
            {
                control = form1.dgv1;
                block = 1;
            }
            else if (callerName == "filePath2" || callerName == "BtnAnalyze2")
            {
                control = form1.dgv2;
                block = 2;
            }
            else
            {
                return;
            }
            control.Rows.Clear();
            var i = 0;
            foreach (var ws in EPPlusHelper.GetExcelWorksheets(excelPackage))
            {
                var index = control.Rows.Add();
                control.Rows[index].Cells[0].Value = ws.Index;
                control.Rows[index].Cells[1].Value = ws.Name;
                var firstValueCellPoint = EPPlusHelper.GetFirstValueCellPoint(ws);
                control.Rows[index].Cells[2].Value = firstValueCellPoint.Row;
                control.Rows[index].Cells[3].Value = firstValueCellPoint.Col;

                if (i == 0 && block == 1)
                {
                    form1.wsNameOrIndex1.Text = control.Rows[index].Cells[1].Value.ToString();
                    form1.TitleLine1.Text = control.Rows[index].Cells[2].Value.ToString();
                    form1.TitleCol1.Text = control.Rows[index].Cells[3].Value.ToString();
                }
                else if (i == 0 && block == 2)
                {
                    form1.wsNameOrIndex2.Text = control.Rows[index].Cells[1].Value.ToString();
                    form1.TitleLine2.Text = control.Rows[index].Cells[2].Value.ToString();
                    form1.TitleCol2.Text = control.Rows[index].Cells[3].Value.ToString();
                }
                i++;
            }
        }

        private void DelHiddenWs_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }
                if (string.Compare(".xlsx", System.IO.Path.GetExtension(filePath), true) != 0)
                {
                    MessageBox.Show("只支持.xlsx文件");
                    return;
                }
                using (var ms = new MemoryStream())
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    EPPlusHelper.DeleteWorksheet(excelPackage, eWorkSheetHidden.Hidden, eWorkSheetHidden.VeryHidden);
                    excelPackage.SaveAs(ms);
                    ms.Position = 0;

                    var fileDir = Path.GetDirectoryName(filePath);
                    var fileName = Path.GetFileNameWithoutExtension(filePath);
                    string filePathOut = Path.Combine(fileDir, $"{fileName}_OnlyVisibleWS.xlsx");
                    ms.Save(filePathOut);
                    MessageBox.Show($"文件已经生成,在目录'{fileDir}'");
                    WinFormHelper.OpenDirectory(fileDir);
                }
            });
        }

        private void TryRun(Action action)
        {
            try
            {
                action.Invoke();
            }
            catch (Exception e)
            {
                MessageBox.Show("程序报错:" + e.Message);
            }
        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            if (dgv.Rows.Count <= 0) return;

            if (e.RowIndex == -1)
            {
                return;//不知道-1 是标格的title 
            }
            var row = dgv.Rows[e.RowIndex];
            var txt = row.Cells[e.ColumnIndex].Value.ToString();

            if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
            {
                if (((System.Windows.Forms.Control)sender).Name == "dgv1") this.wsNameOrIndex1.Text = txt;
                else if (((System.Windows.Forms.Control)sender).Name == "dgv2") this.wsNameOrIndex2.Text = txt;
            }
            else if (e.ColumnIndex == 2)
            {
                if (((System.Windows.Forms.Control)sender).Name == "dgv1")
                    this.TitleLine1.Text = txt;
                else if (((System.Windows.Forms.Control)sender).Name == "dgv2") this.TitleLine2.Text = txt;
            }
            else if (e.ColumnIndex == 3)
            {
                if (((System.Windows.Forms.Control)sender).Name == "dgv1")
                    this.TitleCol1.Text = txt;
                else if (((System.Windows.Forms.Control)sender).Name == "dgv2")
                    this.TitleCol2.Text = txt;
            }
        }

        private void CreateClass_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }
                var fileDir = Path.GetDirectoryName(filePath);
                var dataConfigInfo = new List<ExcelDataConfigInfo>() { GetExcelDataConfigInfo() };

                string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(excelPackage, dataConfigInfo);
                    var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}";
                    var hasFile = false;
                    foreach (var item in defaultConfigList)
                    {
                        if (item.ClassPropertyList.Count > 0)
                        {
                            hasFile = true;
                            File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                        }
                    }
                    if (hasFile)
                    {
                        MessageBox.Show($"文件已经生成,在目录'{fileDir}'");
                    }
                    //if (!filePath.GetDirectoryName().Contains($@"\Desktop\"))
                    //{
                    //    WinFormHelper.OpenFilePath(filePath.GetDirectoryName());
                    //}
                }

            });
        }

        private void CreateDataTable_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show(@"路径不能为空");
                    return;
                }
                var fileDir = Path.GetDirectoryName(filePath);
                var dataConfigInfo = new List<ExcelDataConfigInfo> { GetExcelDataConfigInfo() };

                string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(excelPackage, dataConfigInfo);
                    var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}";

                    var hasFile = false;
                    foreach (var item in defaultConfigList)
                    {
                        if (item.ClassPropertyList.Count > 0)
                        {
                            hasFile = true;
                            File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDataTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDataTableSnippe);
                        }
                    }
                    //if (!filePath.GetDirectoryName().Contains($@"\Desktop\"))
                    //{
                    //    WinFormHelper.OpenFilePath(filePath.GetDirectoryName());
                    //}
                    if (hasFile)
                    {
                        MessageBox.Show($@"文件已经生成,在目录'{fileDir}'");
                    }
                }

            });
        }

        private ExcelDataConfigInfo GetExcelDataConfigInfo()
        {
            var excelDataConfigInfo = new ExcelDataConfigInfo()
            {
                TitleLine = Convert.ToInt32(this.TitleLine1.Text.Trim()),
                TitleColumn = Convert.ToInt32(this.TitleCol1.Text.Trim()),
            };
            var workSheetIndexOrName = this.wsNameOrIndex1.Text.Trim();
            if (int.TryParse(workSheetIndexOrName, out int wsIndex))
            {
                excelDataConfigInfo.WorkSheetIndex = wsIndex;
            }
            else
            {
                excelDataConfigInfo.WorkSheetName = workSheetIndexOrName;
            }

            return excelDataConfigInfo;
        }


        private void diaplayRowAndColumn_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath1.Text.Trim().移除路径前后引号();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("路径不能为空");
                    return;
                }
                string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));

                var filePathNew = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_DisplayAllRows.xlsx";

                using (var ms = new MemoryStream())
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var excelPackage = new ExcelPackage(fs))
                {
                    foreach (var ws in EPPlusHelper.GetExcelWorksheets(excelPackage))
                    {
                        EPPlusHelper.EachHiddenRow(ws, 1, EPPlusConfig.MaxRow07, a => a.Hidden = false);
                        EPPlusHelper.EachHiddenColumn(ws, 1, EPPlusConfig.MaxCol07, a => a.Hidden = false);
                    }
                    excelPackage.SaveAs(ms);
                    ms.Position = 0;
                    ms.Save(filePathNew);
                }

            });
        }
    }
}
