using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    public class DisplayExcelColIndex : System.Attribute
    {
        public int Index { get; set; }

        /// <summary>
        /// 从1开始的
        /// </summary>
        /// <param name="index"></param>
        public DisplayExcelColIndex(int index)
        {
            this.Index = index;
        }
    }
}
