using System.IO;
using System.Windows.Forms;

namespace EPPlusTool.Helper
{
    internal class WinFormHelper
    {

        /// <summary>
        /// 弹出一个选择文件的对话框
        /// </summary>
        /// <returns></returns>
        public static string SelectFile(string filter = null)
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

        public static void OpenFilePath(string savePath)
        {
            if (System.IO.Directory.Exists(savePath))
            {
                //MessageBox.Show($"文件已经生成,在'{savePath}'");
                System.Diagnostics.Process.Start(savePath);
            }
            else
            {
                //MessageBox.Show($"文件已经生成,在'{savePath}'");
                System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(savePath));
            }
        }

        /// <summary>
        /// 打开目录
        /// </summary>
        /// <param name="fileDirectoryName"></param>
        public static void OpenDirectory(string fileDirectoryName)
        {
            //MessageBox.Show($"文件已经生成,在目录'{fileDirectoryName}'");

            //这个会发生异常:System.ComponentModel.Win32Exception:“拒绝访问。”
            //System.Diagnostics.Process.Start(fileDirectoryName); 
            //改用这个
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "explorer.exe";
            process.StartInfo.Arguments = fileDirectoryName;
            //process.StartInfo.FileName = "iexplore.exe";   //IE浏览器，可以更换
            //process.StartInfo.Arguments = "http://www.baidu.com";
            process.Start();
        }

        /// <summary>
        /// 提示文件生成路径并且打开文件所在目录
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="openDirectory"></param>
        public static void PromptFilePathAndOpenDirectory(string filepath, string openDirectory)
        {
            if (!string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show($"文件已经生成,在'{filepath}'");
            }
            if (string.IsNullOrEmpty(openDirectory))
            {
                if (!string.IsNullOrEmpty(filepath))
                {
                    var directoryName = Path.GetDirectoryName(filepath);
                    OpenDirectory(directoryName);
                }
            }
            else
            {
                OpenDirectory(openDirectory);
            }

        }

    }
}
