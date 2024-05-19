using System;

namespace EPPlusTool.Handler
{
    public static class GlobalExceptionHandler
    {
        public static void Register()
        {
            System.Windows.Forms.Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            // 处理异常逻辑
            HandleException(e.Exception);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // 处理异常逻辑
            Exception exception = e.ExceptionObject as Exception;
            HandleException(exception);
        }

        private static void HandleException(Exception exception)
        {

            // 在这里处理异常，例如记录日志或显示错误消息框

#if DEBUG

#else
            var isUnkonwEx = true;//未知异常
            if (exception is BizException)
            {
                MessageBox.Show(exception.Message + ":" + exception.InnerException.Message);
                isUnkonwEx = false;
            }
            try
            {
                // 写入文件
                string filePath = "exception_log.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine(exception.Message);
                    writer.WriteLine("StackTrace:");
                    writer.WriteLine(exception.StackTrace);
                    writer.WriteLine("----------------------------------------");
                }
            }
            catch (global::System.Exception)
            {
            }

            if (isUnkonwEx)
            {
                MessageBox.Show($"未处理的全局异常：{exception.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
#endif

        }
    }
}
