﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EPPlusExtensions;
using EPPlusHelperTool.MethodExtension;

namespace EPPlusHelperTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 弹出一个选择文件的对话框
        /// </summary>
        /// <returns></returns>
        private string SelectFile(string filter = null)
        {
            //显示选择 文件对话框'
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
                //openFileDialog1.InitialDirectory = "c:\\";
                //openFileDialog1.Filter = "excel (*.xlsx)|*.xlsx";
                if (filter != null)
                {
                    openFileDialog1.Filter = filter;
                }
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                var dialogResult = openFileDialog1.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    return openFileDialog1.FileName; //显示文件路径
                }
                else
                {
                    return openFileDialog1.SafeFileName;
                }
            }

        }

        private void TextBoxDragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            ((System.Windows.Forms.TextBox)sender).Text = path;
            if (((System.Windows.Forms.Control)sender).Name == "filePath1")
            {
                WScount1_Click(null, null);
            }
            if (((System.Windows.Forms.Control)sender).Name == "filePath2")
            {
                WScount2_Click(null, null);
            }
        }

        private void TextBoxDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Link : DragDropEffects.None;
        }


        /// <summary>
        /// 打开目录
        /// </summary>
        /// <param name="fileDirectoryName"></param>
        private void OpenDirectory(string fileDirectoryName)
        {
            MessageBox.Show($"文件已经生成,在目录'{fileDirectoryName}'");
            System.Diagnostics.Process.Start(fileDirectoryName);
        }

        private static ExcelWorksheet GetWorkSheet(ExcelPackage excelPackage, string ws1Index_string)
        {
            if (excelPackage.Workbook.Worksheets.Count == 1)
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, 1);
            }
            if (Int32.TryParse(ws1Index_string, out int ws1Index_int))
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, ws1Index_int);
            }
            if (EPPlusHelper.GetExcelWorksheetNames(excelPackage).Contains(ws1Index_string))
            {
                return EPPlusHelper.GetExcelWorksheet(excelPackage, ws1Index_string);
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
                if (this.dgv1.Rows.Count == 0)
                {
                    WScount1_Click(null, null);
                }
                //var fileName = Path.GetFileNameWithoutExtension(filePath);
                //var suffix = Path.GetExtension(filePath);
                var fileDir = Path.GetDirectoryName(filePath);

                //Path.GetDirectoryName(Path.GetFullPath(tempPath))
                //string filePathOut = Path.Combine(fileDir, $"{fileName}_result{suffix}");
                //EPPlusHelper.FillExcelDefaultConfig(filePath, filePathOut);

                var columnTypeList_DateTime = new List<string>()
            {
                "时间", "日期", "date", "time","今天","昨天","明天","前天","day"
            };
                var columnTypeList_String = new List<string>()
            {
                "id","身份证","银行卡","卡号","手机","mobile","tel","序号","number","编号","No"
            };
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

                var dataConfigInfo = new List<ExcelDataConfigInfo>();
                for (int i = 0; i < dgv1.Rows.Count; i++)
                {
                    dataConfigInfo.Add(new ExcelDataConfigInfo()
                    {
                        WorkSheetIndex = i + 1,
                        TitleLine = Convert.ToInt32(dgv1.Rows[i].Cells[2].Value),
                        TitleColumn = Convert.ToInt32(dgv1.Rows[i].Cells[3].Value),
                    });
                }

                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileDir, dataConfigInfo, cell =>
                {
                    var cellValue = EPPlusHelper.GetCellText(cell);
                    var cellValueLower = cellValue.ToLower();
                    foreach (var item in columnTypeList_DateTime)
                    {
                        if (cellValueLower.IndexOf(item, StringComparison.Ordinal) != -1)
                        {
                            cell.Style.Numberformat.Format = "yyyy-mm-dd"; //默认显示的格式
                            break;
                        }
                    }
                    foreach (var item in columnTypeList_DateTime)
                    {
                        if (cellValueLower.IndexOf(item, StringComparison.Ordinal) != -1)
                        {
                            cell.Style.Numberformat.Format = "@"; //Format as text
                            break;
                        }
                    }
                });

                var haveConfig = defaultConfigList.Find(a => a.ClassPropertyList.Count > 0) != null;
                if (!haveConfig)
                {
                    MessageBox.Show("未检测到配置信息");
                    return;
                }

                OpenDirectory(fileDir);
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
                if (this.dgv1.Rows.Count == 0)
                {
                    WScount1_Click(null, null);
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

                //foreach (var item in defaultConfigList.Where(item => item.ClassPropertyList.Count > 0))
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
                WinFormHelper.OpenFilePath(filePath.GetDirectoryName());
                if (errMsg.Length > 0)
                {
                    MessageBox.Show($"下列工作簿未生成配置项:{errMsg}");
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
                if (this.dgv1.Rows.Count == 0)
                {
                    WScount1_Click(null, null);
                }
                if (this.dgv1.Rows.Count == 0)
                {
                    WScount2_Click(null, null);
                }
                var ws1Index_string = this.wsNameOrIndex1.Text.Trim();
                var ws2Index_string = this.wsNameOrIndex2.Text.Trim();

                var ws1TitleLine = Convert.ToInt32(this.TitleLine1.Text.Trim());
                var ws2TitleLine = Convert.ToInt32(this.TitleLine2.Text.Trim());

                var ws1TitleCol = Convert.ToInt32(this.TitleCol1.Text.Trim());
                var ws2TitleCol = Convert.ToInt32(this.TitleCol2.Text.Trim());

                //using (FileStream fs1 = System.IO.File.OpenRead(ws1Path))
                //using (FileStream fs2 = System.IO.File.OpenRead(ws2Path))
                using (FileStream fs1 = new FileStream(ws1Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (FileStream fs2 = new FileStream(ws2Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage1 = new ExcelPackage(fs1))
                using (ExcelPackage excelPackage2 = new ExcelPackage(fs2))
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

                    MessageBox.Show("通过校验模板配置项");
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
            var selectFilePath = SelectFile("excel (*.xlsx)|*.xlsx");
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


        private void WScount1_Click(object sender, EventArgs e)
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
                using (MemoryStream ms = new MemoryStream())
                ////using (FileStream fs = System.IO.File.OpenRead(filePath))
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    var control = this.dgv1;
                    StringBuilder names = new StringBuilder();
                    SetDataSourceForDGV(excelPackage, control, names);

                    if (EPPlusHelper.GetWorkSheetNames(excelPackage, eWorkSheetHidden.Hidden, eWorkSheetHidden.VeryHidden).Count > 0)
                    {
                        MessageBox.Show("当前Excel含有隐藏工作簿,建议删除所有隐藏工作簿");
                    }
                }
            });
        }

        private void WScount2_Click(object sender, EventArgs e)
        {
            TryRun(() =>
            {
                string filePath = filePath2.Text.Trim().移除路径前后引号();
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
                using (MemoryStream ms = new MemoryStream())
                ////using (FileStream fs = System.IO.File.OpenRead(filePath))
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    var control = this.dgv2;
                    StringBuilder names = new StringBuilder();
                    SetDataSourceForDGV(excelPackage, control, names);
                }
            });
        }

        private static void SetDataSourceForDGV(ExcelPackage excelPackage, DataGridView control, StringBuilder names)
        {
            control.Rows.Clear();
            var count = excelPackage.Workbook.Worksheets.Count;

            for (int i = 1; i <= count; i++)
            {
                int index = control.Rows.Add();
                control.Rows[index].Cells[0].Value = i;
                control.Rows[index].Cells[1].Value = excelPackage.Compatibility.IsWorksheets1Based
                    ? excelPackage.Workbook.Worksheets[i].Name
                    : excelPackage.Workbook.Worksheets[i - 1].Name;
                control.Rows[index].Cells[2].Value = 1;
                control.Rows[index].Cells[3].Value = 1;
                names.Append($"{excelPackage.Workbook.Worksheets[i].Name},");
            }

            //var msg = $"一共有{count}个工作簿,分别是:{names.RemoveLastChar(',')}";
            //MessageBox.Show(msg);
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
                using (MemoryStream ms = new MemoryStream())
                ////using (FileStream fs = System.IO.File.OpenRead(filePath))
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    EPPlusHelper.DeleteWorksheet(excelPackage, eWorkSheetHidden.Hidden, eWorkSheetHidden.VeryHidden);
                    excelPackage.SaveAs(ms);
                    ms.Position = 0;

                    var fileDir = Path.GetDirectoryName(filePath);
                    var fileName = Path.GetFileNameWithoutExtension(filePath);
                    string filePathOut = Path.Combine(fileDir, $"{fileName}_OnlyVisibleWS.xlsx");
                    ms.Save(filePathOut);
                    OpenDirectory(fileDir);
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
                else if (((System.Windows.Forms.Control)sender).Name == "dgv2") this.TitleCol2.Text = txt;
            }
        }
    }
}
