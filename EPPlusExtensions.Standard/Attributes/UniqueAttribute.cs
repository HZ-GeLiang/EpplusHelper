using System;

namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// 值唯一
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class UniqueAttribute : Attribute
    {
        public string ErrorMessage { get; private set; }
        public UniqueAttribute() { }

        public UniqueAttribute(string errorMessage)
        {
            this.ErrorMessage = errorMessage;
        }

    }
}
