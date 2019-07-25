using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
            //显示选择 文件对话框
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
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

        private void textBoxDragDrop(object sender, DragEventArgs e)
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

        private void textBoxDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Link : DragDropEffects.None;
        }

        /// <summary>
        /// 打开文件目录
        /// </summary>
        /// <param name="filePath"></param>
        private void OpenFileDirectory(string filePath)
        {
            MessageBox.Show($"文件已经生成,在'{filePath}'");
            System.Diagnostics.Process.Start(Path.GetDirectoryName(filePath));
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
                if (this.dataGridViewExcel1.Rows.Count == 0)
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

                Dictionary<int, int> sheetTitleLineNumber = new Dictionary<int, int>();
                for (int i = 0; i < dataGridViewExcel1.Rows.Count; i++)
                {
                    var titleLine = Convert.ToInt32(dataGridViewExcel1.Rows[i].Cells[2].Value);
                    sheetTitleLineNumber.Add(i, titleLine);
                }

                EPPlusHelper.FillExcelDefaultConfig(filePath, fileDir, sheetTitleLineNumber, cell =>
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
                if (this.dataGridViewExcel1.Rows.Count == 0)
                {
                    WScount1_Click(null, null);
                }

                Dictionary<int, int> sheetTitleLineNumber = new Dictionary<int, int>();
                for (int i = 0; i < dataGridViewExcel1.Rows.Count; i++)
                {
                    var titleLine = Convert.ToInt32(dataGridViewExcel1.Rows[i].Cells[2].Value);
                    sheetTitleLineNumber.Add(i, titleLine);
                }

                string fileOutDirectoryName = Path.GetDirectoryName(Path.GetFullPath(filePath));
                var defaultConfigList = EPPlusHelper.FillExcelDefaultConfig(filePath, fileOutDirectoryName, sheetTitleLineNumber);
                var filePathPrefix = $@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result";
                foreach (var item in defaultConfigList)
                {
                    //将字符串全部写入文件
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateDateTableSnippe)}_{item.WorkSheetName}.txt", item.CrateDateTableSnippe);
                    File.WriteAllText($@"{filePathPrefix}_{nameof(item.CrateClassSnippe)}_{item.WorkSheetName}.txt", item.CrateClassSnippe);
                }
                WinFormHelper.OpenFilePath(filePath.GetDirectoryName());
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
                if (this.dataGridViewExcel1.Rows.Count == 0)
                {
                    WScount1_Click(null, null);
                }
                if (this.dataGridViewExcel1.Rows.Count == 0)
                {
                    WScount2_Click(null, null);
                }
                var ws1Index_string = this.wsNameOrIndex1.Text.Trim();
                var ws2Index_string = this.wsNameOrIndex2.Text.Trim();

                var ws1TitleLine = Convert.ToInt32(this.TitleLine1.Text.Trim());
                var ws2TitleLine = Convert.ToInt32(this.TitleLine2.Text.Trim());

                //using (FileStream fs1 = System.IO.File.OpenRead(ws1Path))
                //using (FileStream fs2 = System.IO.File.OpenRead(ws2Path))
                using (FileStream fs1 = new FileStream(ws1Path, FileMode.Open, FileAccess.Read, FileShare.None))
                using (FileStream fs2 = new FileStream(ws2Path, FileMode.Open, FileAccess.Read, FileShare.None))
                using (ExcelPackage excelPackage1 = new ExcelPackage(fs1))
                using (ExcelPackage excelPackage2 = new ExcelPackage(fs2))
                {
                    var ws1 = GetWorkSheet(excelPackage1, ws1Index_string);
                    var ws2 = GetWorkSheet(excelPackage2, ws2Index_string);
                    var ws1Props = EPPlusHelper.FillExcelDefaultConfig(ws1, ws1TitleLine).ClassPropertyList;
                    var ws2Props = EPPlusHelper.FillExcelDefaultConfig(ws2, ws2TitleLine).ClassPropertyList;

                    StringBuilder sb1 = new StringBuilder();
                    StringBuilder sb2 = new StringBuilder();

                    foreach (var item in ws1Props)
                    {
                        if (!ws2Props.Contains(item))
                        {
                            sb1.Append($@"{item},");
                        }
                    }
                    if (sb1.Length > 1)
                    {
                        MessageBox.Show($@"未提供列:{sb1.RemoveLastChar()}");
                        return;
                    }

                    foreach (var item in ws2Props)
                    {
                        if (!ws1Props.Contains(item))
                        {
                            sb2.Append($@"{item},");
                        }
                    }
                    if (sb2.Length > 1)
                    {
                        MessageBox.Show($@"多提供列:{sb2.RemoveLastChar()}");
                        return;
                    }

                    MessageBox.Show("通过校验模板配置项");
                }
            });
        }

        private void btn_SelectExcelFile(object sender, EventArgs e)
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
                using (MemoryStream ms = new MemoryStream())
                ////using (FileStream fs = System.IO.File.OpenRead(filePath))
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    var control = this.dataGridViewExcel1;
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
                using (MemoryStream ms = new MemoryStream())
                ////using (FileStream fs = System.IO.File.OpenRead(filePath))
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    var control = this.dataGridViewExcel2;
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
                control.Rows[index].Cells[1].Value = excelPackage.Workbook.Worksheets[i].Name;
                control.Rows[index].Cells[2].Value = 1;
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

        private void DataGridViewExcel1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
