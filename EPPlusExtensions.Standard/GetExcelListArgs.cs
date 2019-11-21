using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions.Attributes;

namespace EPPlusExtensions
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T">T是给Filter用的</typeparam>
    public class GetExcelListArgs<T> where T : class
    {
        /// <summary>
        /// excel模板数据从哪列开始
        /// </summary>
        public int DataColStart { get; set; } = 1;

        /// <summary>
        /// excel模板数据从哪列结束
        /// </summary>
        public int? DataColEnd { get; set; } = EPPlusConfig.MaxCol07;

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

        public ScanLine ScanLine = ScanLine.MergeLine;

        public bool MatchingModelEqualsCheck = true;

        /// <summary>
        /// GetList异常时,获得全部异常,而不是一个
        /// </summary>
        public bool GetList_NeedAllException = false;

        /// <summary>
        /// 当GetList_NeedAllException 为 true 时, 错误消息只显示列信息
        /// </summary>
        public bool GetList_ErrorMessage_OnlyShowColomn = false;

        /// <summary>
        /// Key是属性名字,Value是该属性的类型的 KVSource&lt;TKey,Tvalue&gt;
        /// </summary>
        public Dictionary<string, object> KVSource = new Dictionary<string, object>();


    }

    [Flags]
    public enum ReadCellValueOption
    {
        /// <summary>
        /// 无
        /// </summary>
        None = 1,
        /// <summary>
        /// 去空格
        /// </summary>
        Trim = 2,
        /// <summary>
        /// 合并行
        /// </summary>
        MergeLine = 4,
        /// <summary>
        /// 转半角
        /// </summary>
        ToDBC = 8,
    }


    public enum ScanLine
    {
        /*
         * 适合案例:Sample02_1
         */
        /// <summary>
        /// 合并行模式(默认,以眼睛看到的为准)
        /// </summary>
        MergeLine = 1,
        /// <summary>
        /// 逐行读取,
        /// </summary>
        SingleLine = 2,
    }

    internal enum MatchingModel
    {
        /// <summary>
        /// must equal Model=>[model:a,b    excel:a,b]
        /// </summary>
        eq = 1,

        /// <summary>
        /// must greater than Model=>[model:a,b    excel:a,b,c]
        /// </summary>
        gt = 2,

        /// <summary>
        /// must less Than Model=>[model:a,b    excel:a]
        /// </summary>
        lt = 4,

        /// <summary>
        /// must not equal Model=>[model:a,b    excel:a,c] ||  [model:a,b    excel:c,d]
        /// </summary>
        neq = 8
    }
}
