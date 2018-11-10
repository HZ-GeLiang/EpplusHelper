using System;
using System.Reflection;

namespace EpplusExtensions.Helper
{
    public class ReflectionHelper
    {
        public static object[] GetAttributeForProperty<T, TAttribute>(string propertyName)
        {
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));
            var pi = GetProperties<T>();
            return GetProperty(pi, propertyName).GetCustomAttributes(typeof(TAttribute), false);
        }

        public static object[] GetAttributeForProperty<TAttribute>(Type propertyType, string propertyName)
        {
            if (propertyName == null) throw new ArgumentNullException(nameof(propertyName));
            var pi = GetProperties(propertyType);
            return GetProperty(pi, propertyName).GetCustomAttributes(typeof(TAttribute), false);
        }

        public static PropertyInfo[] GetProperties<T>()
        {
            Type type = typeof(T);
            return GetProperties(type);
        }


        public static PropertyInfo[] GetProperties(Type type)
        {
            return type.GetProperties();
        }
        public static PropertyInfo GetProperty(PropertyInfo[] properties, string propertyName)
        {
            return GetProperty(properties, propertyName, false);
        }
        public static PropertyInfo GetProperty(PropertyInfo[] properties, string propertyName, bool notFindReturnNull)
        {
            if (properties == null) throw new ArgumentNullException(nameof(properties));
            foreach (var prop in properties)
            {
                if (prop.Name != propertyName) continue;
                return prop;
            }
            if (notFindReturnNull)
            {
                return null;
            }
            throw new ArgumentOutOfRangeException(nameof(propertyName));
        }
        public static object GetPropertyValue<T>(T model, string propertyName)
        {
            var pi = GetProperties<T>();
            return GetProperty(pi, propertyName).GetValue(model);
        }



    }
}
