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
        public EPPlusConfigSource()
        {
            //SheetHead = new Dictionary<string, string>();
            //SheetBody = new Dictionary<int, DataTable>();
            //SheetBodyFillModel = new Dictionary<int, SheetBodyFillDataMethod>();
            //SheetBodySummary = new Dictionary<int, Dictionary<object, object>>();
            //SheetFoot = new Dictionary<string, string>();
        }
        //public Dictionary<string, string> SheetHead { get; set; }
        //public Dictionary<int, DataTable> SheetBody { get; set; }
        //public Dictionary<int, SheetBodyFillDataMethod> SheetBodyFillModel { get; set; }
        //public Dictionary<int, Dictionary<object, object>> SheetBodySummary { get; set; }
        //public Dictionary<string, string> SheetFoot { get; set; }
        public EPPlusConfigSourceHead Head { get; set; } //= new EPPlusConfigSourceHead();
        public EPPlusConfigSourceBody Body { get; set; } //= new EPPlusConfigSourceBody();
        public EPPlusConfigSourceFoot Foot { get; set; } //= new EPPlusConfigSourceFoot();

    }

    public class EPPlusConfigSourceCellsInfo
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



    public class EPPlusConfigSourceHeadOrFoot
    {
        public List<EPPlusConfigSourceCellsInfo> CellsInfoList { get; set; } = null;
    }
    public class EPPlusConfigSourceHead : EPPlusConfigSourceHeadOrFoot
    {
    }

    public class EPPlusConfigSourceFoot : EPPlusConfigSourceHeadOrFoot
    {
    }

    public class EPPlusConfigSourceBody
    {
        /// <summary>
        /// 所有的配置信息
        /// </summary>
        public List<EPPlusConfigSourceBodyInfo> InfoList { get; set; } = null;
    }

    public class EPPlusConfigSourceBodyInfo
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
        //SheetBody = new Dictionary<int, DataTable>();
        //SheetBodyFillModel = new Dictionary<int, SheetBodyFillDataMethod>();
        //SheetBodySummary = new Dictionary<int, Dictionary<object, object>>();

        public DataTable DataSource { get; set; } = null;
        public SheetBodyFillDataMethod FillMethod { get; set; } = null;
        public List<EPPlusConfigSourceCellsInfo> Summary { get; set; } = null;

    }


}
