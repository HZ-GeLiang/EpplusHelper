using EPPlusExtensions;
using System;
using System.Collections.Generic;

namespace SampleApp._05自动初始化填充配置
{
    public class Sample01
    {
        public static string Run()
        {
            var result = EPPlusHelper.GetFillDefaultConfig("序号	工号	姓名	性别");
            Console.WriteLine(result); //"$tb1序号	$tb1工号	$tb1姓名	$tb1性别";
            return result;
        }
    }

    public class Sample01_alias
    {
        public static string Run()
        {
            var alias = new Dictionary<string, string>()
            {
                {"序号" ,"Index" }
            };
            var result = EPPlusHelper.GetFillDefaultConfig("序号	工号	姓名	性别", alias: alias);
            Console.WriteLine(result); //"$tb1Index	$tb1工号	$tb1姓名	$tb1性别";
            return result;
        }
    }
}