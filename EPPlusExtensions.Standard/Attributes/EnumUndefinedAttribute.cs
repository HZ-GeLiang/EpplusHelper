using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    public class EnumUndefinedAttribute : System.Attribute
    {
        public string ErrorMessage { get; set; }
        public string[] Args { get; set; }
        public EnumUndefinedAttribute(string errorMessageFormat, params string[] args)
        {
            this.ErrorMessage = errorMessageFormat;
            this.Args = args;
        }
    }
}
