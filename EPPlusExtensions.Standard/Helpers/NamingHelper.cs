using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPPlusExtensions.Helpers
{
    /// <summary>
    /// 命名帮助类
    /// </summary>
    internal sealed class NamingHelper
    {
        /// <summary>
        /// 提取符合c#规范的名字
        /// </summary>
        /// <param name="colName"></param>
        /// <returns></returns>
        internal static string ExtractName(string colName)
        {
            string reg = @"[_a-zA-Z\u4e00-\u9FFF][A-Za-z0-9_\u4e00-\u9FFF]*";//去掉不合理的属性命名的字符串(提取合法的字符并接成一个字符串)
            colName = RegexHelper.GetStringByReg(colName, reg).Aggregate("", (current, item) => current + item);
            return colName;
        }

        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        internal static void AutoRename(List<string> nameList, Dictionary<string, int> nameRepeatCounter, string name, bool renameFirstNameWhenRepeat)
        {
            if (!nameRepeatCounter.ContainsKey(name))
            {
                nameRepeatCounter.Add(name, 0);
            }

            if (!nameList.Contains(name) && nameRepeatCounter[name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name] == 1)
                {
                    for (int i = 0; i < nameList.Count; i++)
                    {
                        if (nameList[i] == name)
                        {
                            nameList[i] = nameList[i] + "1";
                            break;
                        }
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name];
                nameList.Add($@"{name}{++currentCounterVal}");
            }

            nameRepeatCounter[name] = ++nameRepeatCounter[name];
        }

        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        internal static void AutoRename(List<ExcelCellInfoValue> nameList, Dictionary<string, int> nameRepeatCounter, ExcelCellInfoValue name, bool renameFirstNameWhenRepeat)
        {
            if (!nameRepeatCounter.ContainsKey(name.Name))
            {
                nameRepeatCounter.Add(name.Name, 0);
            }

            if (nameList.Find(a => a.Name == name.Name) is null && nameRepeatCounter[name.Name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name.Name] == 1)
                {
                    foreach (var t in nameList)
                    {
                        if (t.Name != name.Name)
                        {
                            continue;
                        }

                        t.IsRename = true;
                        t.NameNew = t.Name + "1";
                        break;
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name.Name];
                name.IsRename = true;
                name.NameNew = $@"{name.Name}{++currentCounterVal}";
                nameList.Add(name);
            }
            nameRepeatCounter[name.Name] = ++nameRepeatCounter[name.Name];
        }

        /// <summary>
        /// 自动修复Excel工作表名称,确保获得的Excel工作表名词是正确的
        /// </summary>
        /// <param name="sheetName">需要修复的工作表名称</param>
        /// <param name="replacement">非法字符的代替字符</param>
        /// <returns>修复后的有效工作表名称</returns>
        internal static string FixSheetName(string sheetName, string replacement = "_")
        {
            sheetName = sheetName?.Trim();

            // 如果名称为空，返回默认名称
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return $"Sheet_{DateTime.Now.ToString("yyyyMMddHHmmssfff")}";
            }

            // 移除非法字符: \ / ? * [ ] 或 .
            string fixedName = Regex.Replace(sheetName, @"[\\/\?\*\[\]\.]", replacement ?? "_");

            // 如果首尾字符是单引号，则移除
            while (fixedName.StartsWith("'"))
            {
                fixedName = fixedName.Substring(1);
            }

            while (fixedName.EndsWith("'"))
            {
                fixedName = fixedName.Substring(0, fixedName.Length - 1);
            }

            // 如果修复后名称为空，返回默认名称
            if (string.IsNullOrWhiteSpace(fixedName))
            {
                return $"Sheet_{DateTime.Now.ToString("yyyyMMddHHmmssfff")}";
            }

            // 截断超过31个字符的部分
            if (fixedName.Length > 31)
            {
                fixedName = fixedName.Substring(0, 31);
            }

            return fixedName;
        }

    }
}
