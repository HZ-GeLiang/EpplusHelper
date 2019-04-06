using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPlusHelperTool
{
    public class WinFormHelper
    {

        public static string 获得文件目录地址()
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            return path.SelectedPath;
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

        public static string GetSaveFilePath()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(path);
            string saveFilePath = di.Parent.FullName; //上级目录
            return saveFilePath;
        }
    }
}
