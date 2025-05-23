﻿using System.Collections;
using System.Reflection;

namespace EPPlusExtensions.Helpers
{
    internal sealed class ReflectionHelper
    {
        #region GetAttributeForProperty

        public static object[] GetAttributeForProperty<T, TAttribute>(string propertyName)
        {
            return GetAttributeForProperty<T, TAttribute>(propertyName, false);
        }

        public static object[] GetAttributeForProperty<TAttribute>(Type modelType, string propertyName)
        {
            return GetAttributeForProperty<TAttribute>(modelType, propertyName, false);
        }

        public static object[] GetAttributeForProperty<T, TAttribute>(string propertyName, bool notFindReturnNull)
        {
            if (propertyName is null)
            {
                throw new ArgumentNullException(nameof(propertyName));
            }
            var pi = GetProperties<T>();
            return GetProperty(pi, propertyName, notFindReturnNull)?.GetCustomAttributes(typeof(TAttribute), false);
        }

        public static object[] GetAttributeForProperty<TAttribute>(Type modelType, string propertyName, bool notFindReturnNull)
        {
            if (modelType is null)
            {
                throw new ArgumentNullException(nameof(modelType));
            }
            if (propertyName is null)
            {
                throw new ArgumentNullException(nameof(propertyName));
            }
            var pi = GetProperties(modelType);
            return GetProperty(pi, propertyName, notFindReturnNull)?.GetCustomAttributes(typeof(TAttribute), false);
        }

        #endregion

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
            if (properties is null)
            {
                throw new ArgumentNullException(nameof(properties));
            }
            foreach (var prop in properties)
            {
                if (prop.Name != propertyName)
                {
                    continue;
                }
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

        public static object[] GetMethodParameterDefault(MethodInfo method)
        {
            // MethodInfo method = type.GetMethod(methodName);
            var objArr = new ArrayList();
            var paras = method.GetParameters();
            foreach (ParameterInfo paraInfo in paras)
            {
                if (paraInfo.ParameterType.IsValueType)
                {
                    objArr.Add(0);
                }
                else
                {
                    objArr.Add(null);
                }
            }
            return objArr.ToArray();
        }


        internal static Dictionary<Type, Dictionary<string, PropertyInfo>> _Cache_GetPropertyInfo = new Dictionary<Type, Dictionary<string, PropertyInfo>>();

        internal static PropertyInfo GetPropertyInfo(string propName, Type type)
        {
            if (propName is null)
            {
                throw new ArgumentNullException(nameof(propName));
            }
            if (!_Cache_GetPropertyInfo.ContainsKey(type))
            {
                _Cache_GetPropertyInfo.Add(type, new Dictionary<string, PropertyInfo>());
            }

            var cache_PropertyInfo = _Cache_GetPropertyInfo[type];

            if (!cache_PropertyInfo.ContainsKey(propName))
            {
                var pInfo = type.GetProperty(propName);
                if (pInfo is null) //防御式编程判断
                {
                    throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                }
                cache_PropertyInfo.Add(propName, pInfo);
            }

            return cache_PropertyInfo[propName];
        }
    }
}