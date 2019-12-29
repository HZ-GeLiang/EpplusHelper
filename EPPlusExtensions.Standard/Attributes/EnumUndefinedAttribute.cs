using System;

namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// 枚举值未定义
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class EnumUndefinedAttribute : Attribute
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
