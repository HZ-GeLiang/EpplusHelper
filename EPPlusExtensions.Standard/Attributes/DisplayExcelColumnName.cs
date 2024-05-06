using System;

namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// Excel上的标题列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class DisplayExcelColumnNameAttribute : Attribute
    {
        public string Name { get; private set; }

        public DisplayExcelColumnNameAttribute(string name)
        {
            this.Name = name;
        }
    }
}
