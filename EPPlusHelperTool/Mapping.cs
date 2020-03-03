using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTool
{
    class ExcelInfo
    {
        public string R1C1 { get; set; }
        public int ColumnIndex { get; set; }
        public string Value { get; set; }
    }

    class ModelProp
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public List<string> Attribute { get; set; }
    }
}
