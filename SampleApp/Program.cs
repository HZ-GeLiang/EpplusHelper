using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp
{
    //Func<float, Func<int, float>> happyWater = new Func<float, int, float>((price, number) => number * price).Currying();
    //Func<float, int, float> happyWater2 = new Func<float, int, float>((price, number) => number * price);


    class Program
    {
        static void Main(string[] args1)
        {
            var stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            //SampleApp._01填充数据.Sample01.Run();
            //SampleApp._01填充数据.Sample02.Run();
            //SampleApp._01填充数据.Sample03.Run();
            //SampleApp._01填充数据.Sample04.Run();
            //SampleApp._01填充数据.Sample05.Run();

            //SampleApp._02填充图片.Sample01.Run(); //我也没搞懂怎么使用,计划任务

            //SampleApp._03读取excel内容.Sample01.Run();
            SampleApp._03读取excel内容.Sample02.Run();
            //SampleApp._03读取excel内容.Sample03.Run();
            //SampleApp._03读取excel内容.Sample04.Run();
            //SampleApp._03读取excel内容.Sample05.Run();
            //SampleApp._03读取excel内容.Sample06.Run();
            //SampleApp._03读取excel内容.Sample07.Run();
            //SampleApp._03读取excel内容.Sample08.Run();
            //SampleApp._03读取excel内容.Sample09.Run();
            //SampleApp._03读取excel内容.Sample10.Run();
            //SampleApp._03读取excel内容.Sample11.Run();
            //SampleApp._03读取excel内容.Sample12.Run();
            //SampleApp._03读取excel内容.Sample13.Run();
            //SampleApp._03读取excel内容.Sample14.Run();
            //SampleApp._03读取excel内容.Sample15.Run();

            //_04填充数据与数据源同步-未测试完全,部分可用,不推荐使用.
            //SampleApp._04填充数据与数据源同步.Sample01.Run();
            //SampleApp._04填充数据与数据源同步.Sample02.Run();
            //SampleApp._04填充数据与数据源同步.Sample03.Run();

            //SampleApp._05自动初始化填充配置.Sample01.Run();
            //SampleApp._05自动初始化填充配置.Sample02.Run();
            //SampleApp._05自动初始化填充配置.Sample03.Run();

            stopwatch.Stop();
            Console.WriteLine("runTime 时差:" + stopwatch.Elapsed);
            Console.WriteLine("runTime 毫秒:" + stopwatch.ElapsedMilliseconds);

            Console.ReadKey();
        }
    }


    public static class CurryingExtensions
    {
        //https://mp.weixin.qq.com/s?__biz=MzAxMTMxMDQ3Mw==&mid=2660105542&idx=1&sn=9519dc358cde59e1c6d27773007d5699&chksm=803a59a0b74dd0b6c8a54d3b0967c5bbf7a7c8e92bc3867cd0d099dc08bb8f6e6a4be4c23881&scene=0&xtrack=1&key=0b6f00fa5c3dca5d8719d70beb5e2fecd35b4d8cfb2f28b7c55737d2cb9e2d2b677bb0d6ee198169e333ad0d16dd0c208befe018725150cd96494049cfd155a423dc435f191349d522125d06b3e0fe60&ascene=1&uin=MTgyMTkyNzMwMg%3D%3D&devicetype=Windows+10&version=62060834&lang=zh_CN&pass_ticket=J7b3DfTgb3w9fp7EBZI7udUSW58lTVIRztEd0OMKb6fh%2B0bx100d9R77pES6VeYd

        public static Func<T1, Func<T2, TOutput>> Currying<T1, T2, TOutput>(this Func<T1, T2, TOutput> f) => x => y => f(x, y);
        public static Func<T1, Func<T2, Func<T3, TOutput>>> Currying<T1, T2, T3, TOutput>(this Func<T1, T2, T3, TOutput> f) => x => y => z => f(x, y, z);
    }

    public static class ExpressionTreeHelper
    {
        /// <summary>
        /// Create object.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="objects"></param>
        /// <returns></returns>
        public static T CreateInstance<T>(this Type type, params object[] objects)
        {
            Type[] typeArray = objects.Select(obj => obj.GetType()).ToArray();
            Func<object[], object> deleObj = BuildDeletgateObj(type, typeArray);
            return (T)deleObj(objects);
        }

        /// <summary>
        /// Get a delegate object and use it to generate a entity class.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="typeList"></param>
        /// <returns></returns>
        private static Func<object[], object> BuildDeletgateObj(Type type, Type[] typeList)
        {
            ConstructorInfo constructor = type.GetConstructor(typeList);
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
