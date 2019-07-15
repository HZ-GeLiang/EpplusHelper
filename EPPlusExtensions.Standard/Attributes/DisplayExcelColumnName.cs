using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    public class DisplayExcelColumnName : System.Attribute
    {
        public string Name { get; set; }

        public DisplayExcelColumnName(string name)
        {
            this.Name = name;
        }
    }
}
