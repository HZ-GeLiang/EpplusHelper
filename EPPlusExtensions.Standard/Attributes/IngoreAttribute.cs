namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// 忽略模型的属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class IngoreAttribute : Attribute
    {
        public IngoreAttribute()
        { }
    }
}