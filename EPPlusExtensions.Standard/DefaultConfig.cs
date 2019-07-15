using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 

namespace EPPlusExtensions
{
    public class DefaultConfig
    {
        public string WorkSheetName { get; set; }
        public string CrateDateTableSnippe { get; set; }
        public string CrateClassSnippe { get; set; }
        public List<ExcelCellInfoValue> ClassPropertyList { get; set; }
    }
}
