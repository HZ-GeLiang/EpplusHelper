using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions
{

    /// <summary>
    /// 配置信息的数据源
    /// </summary>
    public class EPPlusConfigSource
    {
        public EPPlusConfigSourceHead Head { get; set; } //= new EPPlusConfigSourceHead();
        public EPPlusConfigSourceBody Body { get; set; } //= new EPPlusConfigSourceBody();
        public EPPlusConfigSourceFoot Foot { get; set; } //= new EPPlusConfigSourceFoot();

    }

    public class EPPlusConfigSourceFixedCell
    {
        /// <summary>
        /// 单元格配置的值:如 Name
        /// </summary>
        public string ConfigValue { get; set; }

        /// <summary>
        /// 填写的值:如 张三
        /// </summary>
        public object FillValue { get; set; }
    }


    public class EPPlusConfigSourceHead  
    {
        public List<EPPlusConfigSourceFixedCell> CellsInfoList { get; set; } = null;
    }

    public class EPPlusConfigSourceFoot  
    {
        public List<EPPlusConfigSourceFixedCell> CellsInfoList { get; set; } = null;
    }

    public class EPPlusConfigSourceBody
    {
        /// <summary>
        /// 所有的配置信息
        /// </summary>
        public List<EPPlusConfigSourceBodyConfig> ConfigList { get; set; } = null;
    }

    public class EPPlusConfigSourceBodyConfig
    {
        /// <summary>
        /// 第几个, 从1开始
        /// </summary>
        public int Nth { get; set; }

        /// <summary>
        /// 对应的设置
        /// </summary>
        public EPPlusConfigSourceBodyOption Option = new EPPlusConfigSourceBodyOption();
    }

    public class EPPlusConfigSourceBodyOption
    {
        /// <summary>
        /// 数据源, 对应  EPPlusConfig 的 ConfigLine
        /// </summary>
        public DataTable DataSource { get; set; } = null;

        /// <summary>
        /// 填充方式
        /// </summary>
        public SheetBodyFillDataMethod FillMethod { get; set; } = null;

        /// <summary>
        /// 固定的一些单元格 如表格的汇总栏什么的
        /// </summary>
        public List<EPPlusConfigSourceFixedCell> ConfigExtra { get; set; } = null;
    }
}
