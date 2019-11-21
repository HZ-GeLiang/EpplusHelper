using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// 枚举值未定义
    /// </summary>
    public class EnumUndefinedAttribute : System.Attribute
    {
        public string ErrorMessage { get; private set; }
        public string[] Args { get; private set; }
        public EnumUndefinedAttribute(string errorMessageFormat, params string[] args)
        {
            this.ErrorMessage = errorMessageFormat;
            this.Args = args;
        }
    }
}
