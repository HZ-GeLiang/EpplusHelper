using EpplusExtensions;
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
            string filePath = textBox1.Text.Trim();

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
    }
}
