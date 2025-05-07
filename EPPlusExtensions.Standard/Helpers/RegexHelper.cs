using System.Text.RegularExpressions;

namespace EPPlusExtensions.Helpers
{
    internal sealed class RegexHelper
    {
        public static List<string> GetStringByReg(string source, string reg)
        {
            MatchCollection regex = Regex.Matches(source, reg);
            List<string> list = new List<string>();
            foreach (Match item in regex)
            {
                list.Add(item.Value);
            }
            return list;
        }

        public static string GetFirstStringByReg(string source, string reg)
        {
            return Regex.Match(source, reg).Groups[0].Value;
        }

        public static string GetFirstNumber(string source)
        {
            return Regex.Match(source, @"\d+").Groups[0].Value;
        }

        public static string GetLastNumber(string source)
        {
            var reg = Regex.Matches(source, @"\d+");
            return reg[reg.Count - 1].Value;
        }
    }
}