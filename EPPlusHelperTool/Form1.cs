using EpplusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPlusHelperTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 弹出一个选择目录的对话框
        /// </summary>
        /// <returns></returns>
        private string SelectPath()
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            return path.SelectedPath;
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
        private void button1_Click(object sender, EventArgs e)
        {
            var selectfilePath = SelectFile("excel (*.xlsx)|*.xlsx");
            if (selectfilePath.Length > 0)
            {
                this.textBox1.Text = selectfilePath;
            }
        }
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            this.textBox1.Text = path;
        }
        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = textBox1.Text.Trim().移除路径前后引号();

            //var fileName = Path.GetFileNameWithoutExtension(filePath);
            //var suffix = Path.GetExtension(filePath);
            var fileDir = Path.GetDirectoryName(filePath);

            //Path.GetDirectoryName(Path.GetFullPath(tempPath))
            //string filePathOut = Path.Combine(fileDir, $"{fileName}_result{suffix}");
            //EpplusHelper.FillExcelDefaultConfig(filePath, filePathOut);
            EpplusHelper.FillExcelDefaultConfig(filePath, fileDir);
            OpenDirectory(fileDir);

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

        private void button3_Click(object sender, EventArgs e)
        {
            var selectfilePath = SelectFile("excel (*.xlsx)|*.xlsx");
            if (selectfilePath.Length > 0)
            {
                this.textBox1.Text = selectfilePath;
            }
        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            this.textBox2.Text = path;
        }
        private void textBox2_DragEnter(object sender, DragEventArgs e)
        {

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var ws1Path = this.textBox1.Text.Trim().移除路径前后引号();
            var ws2Path = this.textBox2.Text.Trim().移除路径前后引号();
            if (ws1Path == ws2Path)
            {
                MessageBox.Show("比较文件路径一致,无法比较");
                return;
            }
            var ws1Index_string = this.textBox3.Text.Trim();
            var ws2Index_string = this.textBox4.Text.Trim();
            var ws1TitleLine = Convert.ToInt32(this.textBox5.Text.Trim());
            var ws2TitleLine = Convert.ToInt32(this.textBox6.Text.Trim());

            using (FileStream fs1 = System.IO.File.OpenRead(ws1Path))
            using (FileStream fs2 = System.IO.File.OpenRead(ws2Path))
            using (ExcelPackage excelPackage1 = new ExcelPackage(fs1))
            using (ExcelPackage excelPackage2 = new ExcelPackage(fs2))
            {
                var ws1 = GetWorkSheet(excelPackage1, ws1Index_string);
                var ws2 = GetWorkSheet(excelPackage2, ws2Index_string);
                var ws1Props = EpplusHelper.FillExcelDefaultConfig(ws1, ws1TitleLine).ClassPropertyList;
                var ws2Props = EpplusHelper.FillExcelDefaultConfig(ws2, ws2TitleLine).ClassPropertyList;

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
        }

        private static ExcelWorksheet GetWorkSheet(ExcelPackage excelPackage, string ws1Index_string)
        {
            if (excelPackage.Workbook.Worksheets.Count == 1)
            {
                return EpplusHelper.GetExcelWorksheet(excelPackage, 1);
            }
            if (Int32.TryParse(ws1Index_string, out int ws1Index_int))
            {
                return EpplusHelper.GetExcelWorksheet(excelPackage, ws1Index_int);
            }
            if (EpplusHelper.GetExcelWorksheetNames(excelPackage).Contains(ws1Index_string))
            {
                return EpplusHelper.GetExcelWorksheet(excelPackage, ws1Index_string);
            }

            throw new ArgumentException("无法打开Excel的Worksheet");
        }
    }
}
