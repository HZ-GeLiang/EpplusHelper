using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions.Attributes;
using EPPlusExtensions.Helper;

namespace EPPlusExtensions
{
    public class GetExcelListArgs
    {

        /// <summary>
        /// excel模板数据从哪列开始,可以理解成标题行的开始列,从1开始
        /// </summary>
        public int DataColStart { get; set; } = 1;

        /// <summary>
        /// excel模板数据从哪列结束,可以理解成标题行的结束列
        /// </summary>
        internal int DataColEnd { get; set; } = EPPlusConfig.MaxCol07;

        /// <summary>
        /// workSheet
        /// </summary>
        public ExcelWorksheet ws { get; set; }

        /// <summary>
        /// 数据的标题行,自带方法提供的默认值是1 即:RowIndex_Data -1
        /// </summary>
        public int DataTitleRow { get; set; }

        /// <summary>
        /// 数据的起始行,自带方法提供的默认值是2
        /// </summary>
        public int DataRowStart { get; set; }

#if DEBUG

        ///// <summary>
        ///// 数据的结尾行,在调用GetList()后自动赋值
        ///// </summary>
        //internal int? DataRowEnd { get; set; }

        ///// <summary>
        ///// 数据有多少行,在调用GetList()后自动赋值
        ///// </summary>
        //internal int? DataRowCount { get; set; }

#endif


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


    }


    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class GetExcelListArgs<T> : GetExcelListArgs where T : class
    {
        //T是给Filter用的,后来也用来创建T类型的Model了

        public T Model { get; set; }

        /// <summary>
        /// 在return数据之前执行过滤操作
        /// </summary>
        public Func<T, bool> HavingFilter { get; set; } = null;

        /// <summary>
        /// 检查数据,如果数据正确,添加到 返回数据 集合中
        /// </summary>
        public Func<T, bool> WhereFilter { get; set; } = null;

    }

    public class KVSource : System.Collections.Generic.Dictionary<string, object> { }

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
         * 适合案例:03读取excel内容 1-3
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

    [Flags]
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
