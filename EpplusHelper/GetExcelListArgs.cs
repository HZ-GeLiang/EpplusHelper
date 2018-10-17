using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpplusExtensions
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T">T是给Filter用的</typeparam>
    public class GetExcelListArgs<T> where T : class, new()
    {
        public ExcelWorksheet ws { get; set; }

        /// <summary>
        /// 数据起始行(不含列名),从1开始
        /// </summary>
        public int rowIndex_Data { get; set; } // = 2;

        /// <summary>
        /// 被遍历的单元格内容不为空时的起始字符必须是该字符,然后忽略该字符
        /// </summary>
        public string EveryCellPrefix { get; set; } = "";

        public Dictionary<string, string> EveryCellReplace { get; set; } = null;

        /// <summary>
        /// EveryCellReplace 的默认提供
        /// </summary>
        public static Dictionary<string, string> EveryCellReplaceDefault = new Dictionary<string, string>
        {
            {"\t", ""},
            {"\r", ""},
            {"\n", ""},
            {"\r\n", ""},
        };

        /// <summary>
        /// 数据起始行(不含列名),从1开始
        /// </summary>
        public int rowIndex_DataName { get; set; }

        public bool UseEveryCellReplace { get; set; } = true;


        /// <summary>
        /// 在return数据之前执行过滤操作
        /// </summary>
        public Func<T, bool> HavingFilter = null;

        /// <summary>
        /// 检查数据,如果数据正确,添加到 返回数据 集合中
        /// </summary>
        public Func<T, bool> WhereFilter = null;

    }
}
