using System;

namespace SampleApp.MethodExtension
{
    internal static class StringExtensions
    {
        public static string RemovePrefix(this string value, string prefix)
        {
            if (value is null) return null;
            if (prefix is null || prefix.Length <= 0) return value;
            return value.StartsWith(prefix) ? value.Substring(prefix.Length, value.Length - prefix.Length) : value;
        }

        /// <summary>
        /// 合并行
        /// </summary>
        /// <param name="src"></param>
        /// <returns></returns>
        public static string MergeLines(this string src)
        {
            if (src is null) throw new ArgumentNullException(nameof(src));
            return src.Replace("\r", "").Replace("\n", "").Replace("\r\n", "");
        }

        public static string RemoveLastChar(this string value, char c)
        {
            return value is null || value.Length <= 0
                ? value
                : (value[value.Length - 1] == c ? value.Remove(value.Length - 1, 1) : value);
        }
    }
}