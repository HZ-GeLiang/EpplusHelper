using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpplusExtensions
{
    public static class StringExtensions
    {
        public static string RemovePrefix(this string value, string prefix)
        {
            if (value == null) return null;
            if (prefix == null || prefix.Length <= 0) return value;
            return value.StartsWith(prefix) ? value.Substring(prefix.Length, value.Length - prefix.Length) : value;
        }

        /// <summary>
        /// 合并行
        /// </summary>
        /// <param name="src"></param>
        /// <returns></returns>
        public static string MergeLines(this string src)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));
            return src.Replace("\r", "").Replace("\n", "").Replace("\r\n", "");
        }

    }
}
