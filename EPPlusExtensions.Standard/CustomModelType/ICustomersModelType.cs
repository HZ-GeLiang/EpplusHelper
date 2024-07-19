using System.Reflection;

namespace EPPlusExtensions.CustomModelType
{
    public interface ICustomersModelType
    {
        /// <summary>
        /// 获得List的时候,有没有Attribute处理
        /// </summary>
        bool HasAttribute { get; set; }

        void RunAttribute<T>(Attribute attribute, PropertyInfo pInfo, T model, string value) where T : class, new();

        void SetModelValue<T>(PropertyInfo pInfo, T model, string value) where T : class, new();
    }
}
