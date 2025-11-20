using System;
using System.IO;
using System.Windows.Forms;

namespace TestOptionFile
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string optionFileName = "options.txt";
            string fullPath = Path.Combine(Application.StartupPath, optionFileName);
            
            Console.WriteLine("当前工作目录: " + Environment.CurrentDirectory);
            Console.WriteLine("应用程序启动路径: " + Application.StartupPath);
            Console.WriteLine("options.txt完整路径: " + fullPath);
            Console.WriteLine("文件是否存在: " + File.Exists(fullPath));
            
            if (File.Exists(fullPath))
            {
                string options = File.ReadAllText(fullPath);
                Console.WriteLine("文件内容: " + options);
                
                var match = System.Text.RegularExpressions.Regex.Match(options, @"<Language>(.*)</Language>");
                if (match.Success)
                {
                    string language = match.Groups[1].Value;
                    Console.WriteLine("语言设置: " + language);
                }
                else
                {
                    Console.WriteLine("未找到语言设置");
                }
            }
            else
            {
                Console.WriteLine("options.txt文件不存在");
            }
        }
    }
}