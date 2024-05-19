using EPPlusTool.Handler;
using System;
using System.Windows.Forms;

namespace EPPlusTool
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            GlobalExceptionHandler.Register(); // 注册全局异常过滤器

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
