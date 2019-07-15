using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{
    public class UniqueAttribute : System.Attribute
    {
        public string ErrorMessage { get; set; }
        public UniqueAttribute() { }

        public UniqueAttribute(string errorMessage)
        {
            this.ErrorMessage = errorMessage;
        }

    }
}
