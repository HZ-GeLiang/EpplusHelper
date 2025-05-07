using EPPlusExtensions.ExtensionMethods;
using EPPlusExtensions.Helpers;
using OfficeOpenXml;
using System.Data;
using System.Text;

namespace EPPlusExtensions
{
    /// <summary>
    /// 程序集内部方法
    /// </summary>
    internal sealed class ExcelAddressHelper
    {
        /// <summary>
        /// 获得属性名
        /// </summary>
        /// <param name="ExcelAddress"></param>
        /// <param name="dictExcelAddressCol"></param>
        /// <param name="dictExcelColumnIndexToModelPropName_All"></param>
        /// <returns>PropName</returns>
        internal static string GetPropName(ExcelAddress ExcelAddress, Dictionary<ExcelAddress, int> dictExcelAddressCol,
            Dictionary<int, string> dictExcelColumnIndexToModelPropName_All)
        {
            int excelCellInfo_ColIndex = dictExcelAddressCol[ExcelAddress];
            if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] is null) //不存在,跳过
            {
                return null;
            }
            return dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
        }

    }
}