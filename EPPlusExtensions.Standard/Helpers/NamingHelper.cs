using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Helpers
{
    /// <summary>
    /// 命名帮助类
    /// </summary>
    public sealed class NamingHelper
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

    }
}
