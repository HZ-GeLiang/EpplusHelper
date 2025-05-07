namespace EPPlusExtensions.ExtensionMethods
{
    internal static class StringExtensions
    {
        public static string RemovePrefix(this string value, string prefix)
        {
            if (value is null)
            {
                return null;
            }

            if (prefix is null || prefix.Length <= 0)
            {
                return value;
            }

            return value.StartsWith(prefix) ? value.Substring(prefix.Length, value.Length - prefix.Length) : value;
        }

        /// <summary>
        /// 合并行
        /// </summary>
        /// <param name="src"></param>
        /// <returns></returns>
        public static string MergeLines(this string src)
        {
            if (src is null)
            {
                throw new ArgumentNullException(nameof(src));
            }

            return src.Replace("\r", "").Replace("\n", "").Replace("\r\n", "");
        }

        public static string RemoveLastChar(this string value, char c)
        {
            return value is null || value.Length <= 0
                ? value
                : value[value.Length - 1] == c ? value.Remove(value.Length - 1, 1) : value;
        }

        /// <summary>
        /// 转半角的函数
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ToDBC(this string str)
        {
            char[] c = str.ToCharArray();
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] == 12288)
                {
                    c[i] = (char)32;
                    continue;
                }

                if (c[i] > 65280 && c[i] < 65375)
                {
                    c[i] = (char)(c[i] - 65248);
                }
            }
            return new string(c);
        }
    }
}