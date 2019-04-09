using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T">T是给Filter用的</typeparam>
    public class GetExcelListArgs<T> where T : class, new()
    {
        public ExcelWorksheet ws { get; set; }


        /// <summary>
        /// 数据起始行(不含列名)
        /// </summary>
        public int RowIndex_Data { get; set; } // = 2;

        /// <summary>
        /// 数据起始行的标题行(不含列名)
        /// </summary>
        public int RowIndex_DataName { get; set; } // RowIndex_Data - 1

        /// <summary>
        /// 被遍历的单元格内容不为空时的起始字符必须是该字符,然后忽略该字符
        /// </summary>
        public string EveryCellPrefix { get; set; } = "";

        /// <summary>
        /// 对每一个单元格进行替换(使用单元格替换)
        /// </summary>
        public bool UseEveryCellReplace { get; set; } = true;

        /// <summary>
        /// 单元格替换列表
        /// </summary>
        public Dictionary<string, string> EveryCellReplaceList { get; set; } = null;

        /// <summary>
        /// EveryCellReplace 的默认提供
        /// </summary>
        public static Dictionary<string, string> EveryCellReplaceListDefault = new Dictionary<string, string>
        {
            {"\t", ""},
            {"\r", ""},
            {"\n", ""},
            {"\r\n", ""},
        };


        /// <summary>
        /// 在return数据之前执行过滤操作
        /// </summary>
        public Func<T, bool> HavingFilter { get; set; } = null;

        /// <summary>
        /// 检查数据,如果数据正确,添加到 返回数据 集合中
        /// </summary>
        public Func<T, bool> WhereFilter { get; set; } = null;

        /// <summary>
        /// 读取每个单元格值时做的处理
        /// </summary>
        public ReadCellValueOption ReadCellValueOption { get; set; } = ReadCellValueOption.Trim;

        /// <summary>
        /// poco属性重名时自动命名
        /// </summary>
        public bool POCO_Property_AutoRename_WhenRepeat { get; set; } = false;

        /// <summary>
        /// poco属性重命名修改第一个名字
        /// </summary>
        public bool POCO_Property_AutoRenameFirtName_WhenRepeat { get; set; } = true;

    }

    public enum ReadCellValueOption
    {
        None = 1,
        Trim = 2,
        MergeLine = 3,
        MergeLineAndTrim = 4,

    }
}
