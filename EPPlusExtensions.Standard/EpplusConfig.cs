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
        /// <summary>
        /// Maximum number of rows in a worksheet (1048576)
        /// </summary>
        public const int MaxRow07 = 1048576;// 和内置的常量值一样 ExcelPackage.MaxRows;
        /// <summary>
        /// Maximum number of columns in a worksheet (16384)
        /// </summary>
        public const int MaxCol07 = 16384; // 和内置的常量值一样 ExcelPackage.MaxColumns;

        //excel 2003 工作表最大有 2^16=65536行,2^8=256列
        //public static readonly int MaxRow03 = 65536;
        //public static readonly int MaxCol03 = 256;

        #endregion

        public EPPlusConfigFixedCells Head { get; set; }

        public EPPlusConfigBody Body { get; set; }

        public EPPlusConfigFixedCells Foot { get; set; }

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
        /// <summary>
        /// 固定单元格信息们
        /// </summary>
        public List<EPPlusConfigFixedCell> ConfigCellList { get; set; } = null;

        /// <summary>
        /// 自定义设置值 action 3个参数 分别代表 (colName,  cellValue, cell)
        /// </summary>
        public Action<string, object, ExcelRange> CellCustomSetValue { get; set; } = null;

        /// <summary>
        /// 填充配置有的列,DataColumn必须存在
        /// </summary>
        public bool ConfigItemMustExistInDataColumn { get; set; } = true;

    }

    /// <summary>
    /// 每一个固定单元格项的配置信息
    /// </summary>
    public class EPPlusConfigFixedCell
    {
        /// <summary>
        /// 单元格地址:如 A8
        /// 不区分大小写,即A2与a2是一样的.建议大写
        /// 对于合并单元格,A2:B2, 返回的也是A2
        /// </summary>
        public string Address { get; set; }

        /// <summary>
        /// 单元格配置的值:如 Name
        /// </summary>
        public string ConfigValue { get; set; }

        /// <summary>
        /// 是否是合并单元格
        /// </summary>
        public bool? IsMergeCell { get; set; }

        /// <summary>
        /// 完整的address,用来给记录合并单元格的
        /// </summary>
        public string FullAddress { get; set; }

    }

    public class EPPlusConfigBody
    {
        /// <summary>
        /// 所有的配置信息
        /// </summary>
        public List<EPPlusConfigBodyConfig> ConfigList { get; set; } = null;


        /// <summary>
        /// 
        /// </summary>
        /// <param name="nth">第几个配置,从1开始</param>
        /// <returns></returns>
        public EPPlusConfigBodyConfig this[int nth]
        {
            get
            {
                if (ConfigList is null) throw new Exception($"{nameof(ConfigList)}为null");
                if (nth < 1) throw new ArgumentOutOfRangeException($"{nameof(nth)}不能小于1");

                var bodyConfig = ConfigList.Find(a => a.Nth == nth);
                if (bodyConfig is null)
                {
                    bodyConfig = new EPPlusConfigBodyConfig()
                    {
                        Nth = nth,
                        Option = new EPPlusConfigBodyOption(),
                    };
                    ConfigList.Add(bodyConfig);
                }
                return bodyConfig;
            }
        }
    }

    public class EPPlusConfigBodyConfig
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

    /// <summary>
    /// 设置的详细内容
    /// </summary>
    public class EPPlusConfigBodyOption
    {
        /// <summary>
        /// body 的内容配置.
        /// </summary>
        public List<EPPlusConfigFixedCell> ConfigLine { get; set; }

        /// <summary>
        /// 配置一共占用了多少列
        /// </summary>
        public int ConfigLineInterval { get; set; }

        /// <summary>
        /// body中固定的单元格. 譬如汇总信息等.譬如A8,Name
        /// </summary>
        public List<EPPlusConfigFixedCell> ConfigExtra { get; set; }

        /// <summary>
        /// 该Action只对ConfigLine有效
        /// 自定义设置值 T1-T4 分别代表 (colName, cellValue, cells, args) 属性名, 属性值, 所在的单元格, 程序内部提供的参数
        /// </summary>
        public Action<CustomValue> CustomSetValue { get; set; }

        /// <summary>
        /// SheetBody模版自带(提供)多少行(根据这个,在结合数据源,程序内部判断是否新增行)
        /// </summary>
        public int? MapperExcelTemplateLine { get; set; }

        /// <summary>
        /// 自定义设置值  
        /// </summary>
        public Action<CustomValue> SummaryCustomSetValue { get; set; }

        public InsertRowStyle InsertRowStyle { get; set; } = new InsertRowStyle();

        /// <summary>
        /// 填充配置有的列,DataColumn必须存在
        /// </summary>
        public bool ConfigItemMustExistInDataColumn { get; set; } = true;
    }

    public class CustomValue
    {

        /// <summary>
        /// 属性名:列名 (要填充的列,如姓名)
        /// </summary>
        public string ColName { get; set; }
        /// <summary>
        ///属性值: 要填充的值 
        /// </summary>
        public object Value { get; set; }
        /// <summary>
        /// 重所在的单元格:填充的单元格
        /// </summary>
        public ExcelRange Cell { get; set; }

        #region 程序内部提供的参数

        /// <summary>
        /// 程序内部提供的参数:配置行信息
        /// </summary>
        public List<EPPlusConfigFixedCell> ConfigLine { get; set; }

        /// <summary>
        /// 程序内部提供的参数:额外要填充的单元格信息
        /// </summary>
        public List<EPPlusConfigFixedCell> ConfigExtra { get; set; }

        /// <summary>
        /// 程序内部提供的参数:填充的工作簿
        /// </summary>
        public ExcelWorksheet Worksheet { get; set; }

        /// <summary>
        /// 程序内部提供的参数:填充的区域
        /// </summary>
        public FillArea? Area { get; set; }

        #endregion
    }

    //public class CustomSetValueArgument
    //{
    //    public List<EPPlusConfigFixedCell> ConfigLine { get; set; }
    //    public List<EPPlusConfigFixedCell> ConfigExtra { get; set; }
    //    public ExcelWorksheet Worksheet { get; set; }
    //    public FillArea Area { get; set; }
    //}

    public enum FillArea
    {
        /// <summary>
        /// 标题
        /// </summary>
        TitleExt = 1,
        /// <summary>
        /// 内容(配置的哪些)
        /// </summary>
        Content = 2,
        /// <summary>
        /// 内容扩展,DataTable 未配置的列)
        /// </summary>
        ContentExt = 3,

    }

    public class InsertRowStyle
    {
        /// <summary>
        /// 插入新增行的操作
        /// </summary>
        public InsertRowStyleOperation Operation { get; set; } = InsertRowStyleOperation.CopyAll;

        #region 这2个是 CopyStyleAndMergedCellFromConfigRow 的配置
        /// <summary>
        /// 新增行时复制配置项所在行的样式(新增的行不含单元格合并) ,相同的工作簿,该选项 false 时, 生成的文件体积会减小很多
        /// </summary>
        public bool NeedCopyStyles { get; set; } = true;

        /// <summary>
        /// 配置行有合并单元格时,新增行也需要
        /// </summary>
        public bool NeedMergeCell { get; set; } = true;
        #endregion

    }

    public enum InsertRowStyleOperation
    {
        /// <summary>
        /// 复制配置行的所有样式(含合并单元格)
        /// </summary>
        CopyAll = 1,
        /// <summary>
        /// 复制配置行的样式,然后合并单元格(如果配置行有)
        /// </summary>
        CopyStyleAndMergeCell = 2
    }
}
