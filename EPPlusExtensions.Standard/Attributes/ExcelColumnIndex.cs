using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    public class ExcelColumnIndexAttribute : System.Attribute
    {
        public int Index { get; set; }

        /// <summary>
        /// 从1开始的
        /// </summary>
        /// <param name="index"></param>
        public ExcelColumnIndexAttribute(int index)
        {
            this.Index = index;
        }
    }
}
