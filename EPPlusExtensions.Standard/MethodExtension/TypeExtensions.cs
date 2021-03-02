using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusExtensions.MethodExtension
{
    internal static class TypeExtensions
    {
        private static Dictionary<string, bool> _cache = new Dictionary<string, bool>();
        //代码来自https://www.cnblogs.com/walterlv/p/10236419.html#NET__8
        /// <summary>
        /// 判断指定的类型 <paramref name="type"/> 是否是指定泛型类型的子类型，或实现了指定泛型接口。
        /// </summary>
        /// <param name="type">需要测试的类型。</param>
        /// <param name="generic">泛型接口类型，传入 typeof(IXxx&lt;&gt;)</param>
        /// <returns>如果是泛型接口的子类型，则返回 true，否则返回 false。</returns>
        //public static bool HasImplementedRawGeneric([NotNull] this Type type, [NotNull] Type generic)
        public static bool HasImplementedRawGeneric(this Type type, Type generic)
        {
            if (type == null) throw new ArgumentNullException(nameof(type));
            if (generic == null) throw new ArgumentNullException(nameof(generic));

            //if (_cache.Keys.Count > 1000)
            //{
            //    _cache.Clear();
            //}

            //var key =$@"{type.Assembly}_{type.AssemblyQualifiedName}|{generic.Assembly}_{generic.AssemblyQualifiedName}";
            var key = $@"{type.GetHashCode()}|{generic.GetHashCode()}";

            if (_cache.ContainsKey(key))
            {
                return _cache[key];
            }

            // 测试接口。
            var isTheRawGenericType = type.GetInterfaces().Any(IsTheRawGenericType);
            if (isTheRawGenericType)
            {
                _cache[key] = true;
                return true;
            }

            // 测试类型。
            while (type != null && type != typeof(object))
            {
                isTheRawGenericType = IsTheRawGenericType(type);
                if (isTheRawGenericType)
                {
                    _cache[key] = true;
                    return true;
                }
                type = type.BaseType;
            }

            // 没有找到任何匹配的接口或类型。
            _cache[key] = false;
            return false;

            // 测试某个类型是否是指定的原始接口。
            bool IsTheRawGenericType(Type test)
                => generic == (test.IsGenericType ? test.GetGenericTypeDefinition() : test);
        }
    }
}
