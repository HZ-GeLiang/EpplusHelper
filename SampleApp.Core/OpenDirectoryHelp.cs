using System;

namespace SampleApp.Core
{
    class OpenDirectoryHelp
    {
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
