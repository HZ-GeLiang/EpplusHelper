using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions.MethodExtension;

namespace EPPlusExtensions.Helper
{
    internal static class ExpressionTreeExtensions
    {
        /// <summary>
        /// Create object.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="objects"></param>
        /// <returns></returns>
        public static T CreateInstance<T>(this Type type, params object[] objects)
        {
            Type[] typeArray;
            if (objects == null || objects.Length == 0)
            {
                typeArray = new Type[0];
            }
            else
            {
                typeArray = objects.Select(obj => obj.GetType()).ToArray();
            }
            Func<object[], object> deleObj = BuildDeletgateCreateInstance(type, typeArray);
            return (T)deleObj(objects);
        }

        /// <summary>
        /// Get a delegate object and use it to generate a entity class.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="typeList"></param>
        /// <returns></returns>
        public static Func<object[], object> BuildDeletgateCreateInstance(Type type, Type[] typeList)
        {
            ConstructorInfo constructor = type.GetConstructor(typeList);
            if (constructor == null)
            {
                if (typeList == null || typeList.Length == 0)
                {
                    throw new Exception($@"未找到类'{type.Name}({type.FullName})'的无参数构造器.");
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in typeList)
                    {
                        sb.Append($@"{item.Name}({item.FullName}),");
                    }

                    sb.RemoveLastChar(',');
                    var msg = $@"未找到类'{type.Name}({type.FullName})'的构造器:参数个数:{typeList.Length},参数类型:{sb}.";
                    throw new Exception(msg);
                }
            }
            ParameterExpression paramExp = Expression.Parameter(typeof(object[]), "args_");
            Expression[] expList = GetExpressionArray(typeList, paramExp);

            NewExpression newExp = Expression.New(constructor, expList);

            Expression<Func<object[], object>> expObj = Expression.Lambda<Func<object[], object>>(newExp, paramExp);
            return expObj.Compile();
        }

        /// <summary>
        /// Get an expression array.
        /// </summary>
        /// <param name="typeList"></param>
        /// <param name="paramExp"></param>
        /// <returns></returns>
        private static Expression[] GetExpressionArray(Type[] typeList, ParameterExpression paramExp)
        {
            List<Expression> expList = new List<Expression>();
            for (int i = 0; i < typeList.Length; i++)
            {
                var paramObj = Expression.ArrayIndex(paramExp, Expression.Constant(i));
                var expObj = Expression.Convert(paramObj, typeList[i]);
                expList.Add(expObj);
            }

            return expList.ToArray();
        }
    }
}
