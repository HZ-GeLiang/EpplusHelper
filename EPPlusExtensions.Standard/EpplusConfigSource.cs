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
            SheetHead = new Dictionary<string, string>();
            SheetBody = new Dictionary<int, DataTable>();
            SheetBodyFillModel = new Dictionary<int, SheetBodyFillDataMethod>();
            SheetBodySummary = new Dictionary<int, Dictionary<object, object>>();
            SheetFoot = new Dictionary<string, string>();
        }
        public Dictionary<string, string> SheetHead { get; set; }
        public Dictionary<int, DataTable> SheetBody { get; set; }
        public Dictionary<int, SheetBodyFillDataMethod> SheetBodyFillModel { get; set; }
        public Dictionary<int, Dictionary<object, object>> SheetBodySummary { get; set; }
        public Dictionary<string, string> SheetFoot { get; set; }

    }
}
