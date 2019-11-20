using System.IO;

namespace EPPlusTool
{
    internal static class StringExtensions
    {
        public static string GetDirectoryName(this string filePath)
        {
            return Path.GetDirectoryName(移除路径前后引号(filePath));
        }
        public static string 移除路径前后引号(this string filePath)
        {
            return filePath != null && filePath.Length > 2 && filePath[0] == '"' && filePath[filePath.Length - 1] == '"'
                ? filePath.Substring(1, filePath.Length - 1 - 1)
                : filePath;
        }
    }
}
