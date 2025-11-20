using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using System.IO;

namespace BulkPDF
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            try
            {
                // 添加调试日志
                File.AppendAllText("debug.log", $"[{DateTime.Now}] 程序启动\n");
                
                // Language
                string option = OptionFileHandler.GetOptionValue("Language");
                File.AppendAllText("debug.log", $"[{DateTime.Now}] 语言选项: {option ?? "null"}\n");
                
                if (!String.IsNullOrEmpty(option))
                    foreach (var cultureInfo in CultureInfo.GetCultures(CultureTypes.AllCultures))
                        if (cultureInfo.Name == option)
                            Thread.CurrentThread.CurrentUICulture = new CultureInfo(cultureInfo.Name);

                // Programm
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                File.AppendAllText("debug.log", $"[{DateTime.Now}] 创建主窗体\n");
                Application.Run(new MainForm());
                File.AppendAllText("debug.log", $"[{DateTime.Now}] 程序正常退出\n");
            }
            catch (Exception ex)
            {
                string errorMsg = $"[{DateTime.Now}] 程序异常: {ex.Message}\n{ex.StackTrace}\n";
                File.AppendAllText("debug.log", errorMsg);
                MessageBox.Show($"程序启动失败:\n{ex.Message}\n\n详细信息已保存到 debug.log 文件", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}