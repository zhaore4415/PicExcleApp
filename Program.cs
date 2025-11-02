using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PicExcleApp;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // 设置全局异常处理
        Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
        AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        
        // 确保所有任务异常都被捕获
        TaskScheduler.UnobservedTaskException += new EventHandler<UnobservedTaskExceptionEventArgs>(TaskScheduler_UnobservedTaskException);
        
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }
    
    // 处理UI线程异常
    private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
    {
        LogException(e.Exception, "UI线程异常");
    }
    
    // 处理非UI线程异常
    private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            LogException(ex, "未处理的异常");
        }
    }
    
    // 处理任务异常
    private static void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
    {
        LogException(e.Exception, "任务异常");
        e.SetObserved(); // 标记异常为已观察
    }
    
    // 记录异常到文件
    private static void LogException(Exception ex, string type)
    {
        try
        {
            string logFilePath = Path.Combine(Application.StartupPath, "error_log.txt");
            string logContent = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {type}\n" +
                               $"错误信息: {ex.Message}\n" +
                               $"错误类型: {ex.GetType().FullName}\n" +
                               $"堆栈跟踪: {ex.StackTrace}\n";
            
            if (ex.InnerException != null)
            {
                logContent += $"内部异常: {ex.InnerException.Message}\n" +
                             $"内部异常堆栈: {ex.InnerException.StackTrace}\n";
            }
            
            logContent += "=============================\n";
            
            // 写入日志文件
            File.AppendAllText(logFilePath, logContent);
            
            // 显示详细错误信息（仅在调试时）
            MessageBox.Show($"发生错误: {ex.Message}\n\n详细信息已记录到错误日志文件。\n\n错误类型: {ex.GetType().Name}\n\n错误日志: {logFilePath}", 
                           "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch (Exception)
        {
            // 如果无法写入日志文件，至少显示基本错误
            MessageBox.Show($"发生错误且无法记录日志: {ex.Message}", 
                           "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }    
}