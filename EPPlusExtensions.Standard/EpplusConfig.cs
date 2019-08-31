using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace EPPlusExtensions
{

    /// <summary>
    /// 配置信息 
    /// </summary>
    public class EPPlusConfig
    {
        #region Excel的最大行与列 

        //excel 2007 和 excel 2010 工作表最大有 2^20=1048576行,2^14=16384列
        public static readonly int MaxRow07 = 1048576;
        public static readonly int MaxCol07 = 16384;

        //excel 2003 工作表最大有 2^16=65536行,2^8=256列
        //public static readonly int MaxRow03 = 65536;
        //public static readonly int MaxCol03 = 256;

        #endregion

        /// <summary>
        /// 用来初始化的一些数据的
        /// </summary>
        public EPPlusConfig()
        {
            //SheetHeadMapperExcel = new Dictionary<string, string>();
            //SheetHeadMapperSource = new Dictionary<string, string>();
            //SheetHeadCellCustomSetValue = null;

            //注:body是没有数据源的配置的,全靠一个默认约定
            //SheetBodyMapperExcel = new Dictionary<int, Dictionary<string, string>>();
            //SheetBodyCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
            //SheetBodySummaryMapperExcel = new Dictionary<int, Dictionary<string, string>>();
            //SheetBodySummaryMapperSource = new Dictionary<int, Dictionary<string, string>>();
            //SheetBodySummaryCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
            //SheetBodyMapperExcelTemplateLine = new Dictionary<int, int>();

            //SheetFootMapperExcel = new Dictionary<string, string>();
            //SheetFootMapperSource = new Dictionary<string, string>();
            //SheetFootCellCustomSetValue = null;

            //Report = new EPPlusReport();
            //IsReport = false;
            //DeleteFillDateStartLineWhenDataSourceEmpty = false;
        }

        #region head

        #region 注释的代码
        ///// <summary>
        ///// sheet head 用来完成指定单元格的内容配置
        ///// 譬如A2,Name. key不区分大小写,即A2与a2是一样的.建议大写
        ///// </summary>
        //public Dictionary<string, string> SheetHeadMapperExcel { get; set; }

        ///// <summary>
        ///// sheet head 的数据源的配置
        ///// 譬如Name,张三. key严格区分大小写
        ///// </summary>
        //public Dictionary<string, string> SheetHeadMapperSource { get; set; }

        ///// <summary>
        ///// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        ///// </summary>
        //public Action<string, object, ExcelRange> SheetHeadCellCustomSetValue; 
        #endregion

        public EPPlusConfigFixedCells Head { get; set; }

        #endregion

        #region body

        #region 注释的代码
        ///// <summary>
        ///// sheet body 的内容配置.注.int必须是从1开始的且递增+1的自然数
        ///// </summary>
        //public Dictionary<int, Dictionary<string, string>> SheetBodyMapperExcel;

        ///// <summary>
        ///// 自定义设置值  action 3个参数 分别代表 (colName,  cellValue, cell)
        ///// </summary>
        //public Dictionary<int, Action<string, object, ExcelRange>> SheetBodyCellCustomSetValue;

        ///// <summary>
        ///// sheet body中固定的单元格. 譬如汇总信息等.譬如A8,Name,前面的int表示这个汇总是哪个SheetBody的
        ///// </summary>
        //public Dictionary<int, Dictionary<string, string>> SheetBodySummaryMapperExcel { get; set; }

        ///// <summary>
        ///// sheet body中固定的单元格的数据源,譬如Name,张三
        ///// </summary>
        //public Dictionary<int, Dictionary<string, string>> SheetBodySummaryMapperSource { get; set; }
        ///// <summary>
        ///// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        ///// </summary>
        //public Dictionary<int, Action<string, object, ExcelRange>> SheetBodySummaryCellCustomSetValue;

        ///// <summary>
        ///// SheetBody模版自带(提供)多少行(根据这个,在结合数据源,程序内部判断是否新增行)
        ///// </summary>
        //public Dictionary<int, int> SheetBodyMapperExcelTemplateLine { get; set; } 
        #endregion

        public EPPlusConfigBody Body { get; set; }

        #endregion

        #region foot

        #region 注释的代码
        ///// <summary>
        ///// sheet foot 用来完成指定单元格的内容配置
        ///// 譬如A8,Name
        ///// </summary>
        //public Dictionary<string, string> SheetFootMapperExcel { get; set; }

        ///// <summary>
        ///// sheet foot 的数据源
        ///// 譬如 Name,张三</summary>
        //public Dictionary<string, string> SheetFootMapperSource { get; set; }

        ///// <summary>
        ///// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        ///// </summary>
        //public Action<string, object, ExcelRange> SheetFootCellCustomSetValue;

        //使用方式
        //config.SheetBodyCellCustomSetValue.Add(1, (colName, cellValue, cell) =>
        //{
        //    if (colName == "配置列名")
        //    {
        //        cell.Formula = (string)val;//值为公式
        //    }
        //    else
        //    {
        //        cell.Value = val;
        //    }
        //} 
        #endregion

        public EPPlusConfigFixedCells Foot { get; set; }

        #endregion

        /// <summary>
        /// 报表(excel能折叠的那种)的显示的一些配置
        /// </summary>
        public EPPlusReport Report { get; set; }

        /// <summary>
        /// 标识是否是一个报表格式(excel能折叠的)的Worksheet(目前该属性表示每一个worksheet), 默认False
        /// </summary>
        public bool IsReport { get; set; }

        /// <summary>
        /// 当填充的数据源为空时,是否删除填充的起始行,默认false
        /// </summary>
        public bool DeleteFillDateStartLineWhenDataSourceEmpty;

        /// <summary>
        /// 是否使用默认(单元格格式)约定,默认true 注:settingCellFormat 若与默认的发成冲突,会把默认的 cell 格式给覆盖.
        /// </summary>
        public bool UseFundamentals = true;
        /// <summary>
        /// 默认的单元格格式设置,colMapperName 是配置单元格的名字 譬如 $tb1Id, 那么colMapperName值就为Id
        /// </summary>

        public Func<string, object, ExcelRange, object> CellFormatDefault = (colMapperName, val, cells) =>
        {
            //关于格式,可以参考右键单元格->设置单元格格式->自定义中的类型 或看文档: https://support.office.microsoft.com/zh-CN/excel ->自定义 Excel->创建或删除自定义数字格式
            string formatStr = cells.Style.Numberformat.Format;
            //含有Id的列,默认是文本类型,目的是放防止出现科学计数法
            if (colMapperName != null && colMapperName.ToLower().EndsWith("id"))
            {
                if (formatStr != "@")
                {
                    cells.Style.Numberformat.Format = "@"; //Format as text
                }
                if (val.GetType() != typeof(string))
                {
                    val = val.ToString(); //确保值是string类型的
                }
            }
            //若没有设置日期格式,默认是yyyy-mm-dd
            //大写字母是为了冗错.注:excel的日期格式写成大写的是不会报错的,但文档中全是小写的.
            var dateCode = new List<char> { '@', 'y', 'Y', 's', 'S', 'm', 'M', 'h', 'H', 'd', 'D', 'A', 'P', ':', '.', '0', '[', ']' };
            if (val is DateTime)
            {
                var changeFormat = true;
                foreach (var c in formatStr) //这边不能用优化成linq,优化成linq有问题
                {
                    if (dateCode.Contains(c))
                    {
                        changeFormat = false;
                        break;
                    }
                }
                if (changeFormat) //若为true,表示没有人为的设置该cell的日期显示格式
                {
                    cells.Style.Numberformat.Format = "yyyy-mm-dd"; //默认显示的格式
                }
            }

            return val;
        };

        public Action<ExcelWorksheet> WorkSheetDefault;
        //= worksheet =>
        //{
        //    //worksheet.DefaultColWidth = 72; //默认列宽
        //    //worksheet.DefaultRowHeight = 18; //默认行高
        //    //worksheet.TabColor = Color.Blue; //Sheet Tab的颜色
        //    //worksheet.Cells.Style.WrapText = true; //单元格文字自动换行
        //};

    }
    /// <summary>
    /// 配置信息-固定的单元格
    /// </summary>
    public class EPPlusConfigFixedCells
    {
        ///// <summary>
        ///// 固定单元格信息们
        ///// </summary>
        public List<EPPlusConfigCellsInfo> ConfigCellList { get; set; } = null;

        ///// <summary>
        ///// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        ///// </summary>
        public Action<string, object, ExcelRange> CellCustomSetValue { get; set; } = null;

    }

    /// <summary>
    /// 每一个固定单元格配置信息
    /// </summary>
    public class EPPlusConfigCellsInfo
    {
        /// <summary>
        /// 单元格地址:如 A8 ,不区分大小写,即A2与a2是一样的.建议大写
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        /// 单元格配置的值:如 Name
        /// </summary>
        public string ConfigValue { get; set; }

        ///// <summary>
        ///// 填写的值:如 张三
        ///// </summary>
        //public string FillValue { get; set; }
    }

    public class EPPlusConfigBody
    {
        /// <summary>
        /// 所有的配置信息
        /// </summary>
        public List<EPPlusConfigBodyInfo> InfoList { get; set; } = null;
    }
    public class EPPlusConfigBodyInfo
    {

        /// <summary>
        /// 第几个配置, 从1开始
        /// </summary>
        public int Nth { get; set; }

        /// <summary>
        /// 对应的设置
        /// </summary>
        public EPPlusConfigBodyOption Option = new EPPlusConfigBodyOption();
    }

    public class EPPlusConfigBodyOption
    {
        //SheetBodyMapperExcel = new Dictionary<int, Dictionary<string, string>>();
        //SheetBodyCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
        //SheetBodySummaryMapperExcel = new Dictionary<int, Dictionary<string, string>>();
        //SheetBodySummaryMapperSource = new Dictionary<int, Dictionary<string, string>>();
        //SheetBodySummaryCellCustomSetValue = new Dictionary<int, Action<string, object, ExcelRange>>();
        //SheetBodyMapperExcelTemplateLine = new Dictionary<int, int>();

        /// <summary>
        /// sheet body 的内容配置.
        /// </summary>
        public List<EPPlusConfigCellsInfo> MapperExcel { get; set; }

        /// <summary>
        /// 自定义设置值  action 3个参数 分别代表 (colName,  cellValue, cell)
        /// </summary>
        public Action<string, object, ExcelRange> CustomSetValue { get; set; }


        ///// <summary>
        ///// sheet body中固定的单元格. 譬如汇总信息等.譬如A8,Name
        ///// </summary>
        public List<EPPlusConfigCellsInfo> SummaryMapperExcel { get; set; }

        /// <summary>
        /// sheet body中固定的单元格的数据源,譬如Name,张三
        /// </summary>
        // todo:...
        public Dictionary<string, string> SummaryMapperSource { get; set; }

        /// <summary>
        /// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        /// </summary>
        public Action<string, object, ExcelRange> SummaryCustomSetValue { get; set; }

        /// <summary>
        /// SheetBody模版自带(提供)多少行(根据这个,在结合数据源,程序内部判断是否新增行)
        /// </summary>
        public int? MapperExcelTemplateLine { get; set; }

    }


}
