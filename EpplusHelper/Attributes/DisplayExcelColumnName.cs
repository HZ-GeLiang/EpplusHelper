using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpplusExtensions.Attributes
{
    public class DisplayExcelColumnNameAttribute : System.Attribute
    {
        public string Name { get; set; }

        public DisplayExcelColumnNameAttribute(string name)
        {
            this.Name = name;
        }

    }
}
