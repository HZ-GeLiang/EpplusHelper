namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// Excel的标题列所在的列序号是多少,从1开始
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelColumnIndexAttribute : Attribute
    {
        public int Index { get; private set; }

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