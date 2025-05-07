using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
namespace EPPlusExtensions.Validators
{
    public class ExcelSheetNameValidator
    {
        private readonly string _sheetName;

        /// <summary>
        ///
        /// </summary>
        /// <param name="sheetName">要检查的工作表名称</param>
        public ExcelSheetNameValidator(string sheetName)
        {
            _sheetName = sheetName?.Trim();
        }

        /// <summary>
        /// 检查Excel工作表名称是否有效
        /// </summary>
        /// <returns>如果名称有效返回true，否则返回false</returns>
        public bool IsValidSheetName()
        {
            string sheetName = _sheetName;

            // 检查是否为空
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            // 检查长度是否超过31个字符
            if (sheetName.Length > 31)
            {
                return false;
            }

            // 检查是否包含非法字符: \ / ? * [ ] 或 .
            if (Regex.IsMatch(sheetName, @"[\\/\?\*\[\]\.]"))
            {
                return false;
            }

            // 检查首尾字符是否为单引号
            if (sheetName.StartsWith("'") || sheetName.EndsWith("'"))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获取工作表名称无效的原因
        /// </summary>
        /// <returns>如果名称有效返回空字符串，否则返回无效原因</returns>
        public string GetInvalidReason()
        {
            string sheetName = _sheetName;

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return "工作表名称不能为空。";
            }

            if (sheetName.Length > 31)
            {
                return "工作表名称不能超过31个字符。";
            }

            if (Regex.IsMatch(sheetName, @"[\\/\?\*\[\]\.]"))
            {
                return "工作表名称中不能包含以下字符：: \\ / ? * [ ] 或 .";
            }

            if (sheetName.StartsWith("'") || sheetName.EndsWith("'"))
            {
                return "工作表名称的第一个或最后一个字符不能是单引号。";
            }

            return string.Empty;
        }


        /// <summary>
        /// 自动修复Excel工作表名称,确保获得的Excel工作表名词是正确的
        /// </summary>
        /// <param name="replacement">非法字符的代替字符</param>
        /// <returns>修复后的有效工作表名称</returns>
        public string GetFixSheetName(string replacement = "_")
        {
            if (IsValidSheetName() == true)
            {
                return _sheetName;
            }

            string sheetName = _sheetName?.Trim();//需要修复的工作表名称

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
