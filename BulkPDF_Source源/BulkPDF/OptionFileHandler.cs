using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace BulkPDF
{
    internal static class OptionFileHandler
    {
        private static string optionFileName = "options.txt";

        public static string GetOptionValue(string name)
        {
            // 使用完整路径查找options.txt文件
            string fullPath = Path.Combine(Application.StartupPath, optionFileName);
            
            if (File.Exists(fullPath))
            {
                string options = File.ReadAllText(fullPath);
                return Regex.Match(options, @"<" + name + @">(.*)</" + name + @">").Groups[1].Value;
            }
            else
            {
                return "";
            }
        }

        public static void SetOptionValue(string name, string value)
        {
            // 使用完整路径查找options.txt文件
            string fullPath = Path.Combine(Application.StartupPath, optionFileName);
            string content = "";
            
            if (File.Exists(fullPath))
            {
                content = File.ReadAllText(fullPath);
                // 替换现有值
                string pattern = @"<" + name + @">.*</" + name + @">";
                string replacement = "<" + name + ">" + value + "</" + name + ">";
                content = Regex.Replace(content, pattern, replacement);
            }
            else
            {
                // 创建新文件
                content = "<" + name + ">" + value + "</" + name + ">";
            }
            
            File.WriteAllText(fullPath, content);
        }
    }
}