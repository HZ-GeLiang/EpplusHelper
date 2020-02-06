using EPPlusExtensions;
using System;

namespace SampleApp._05自动初始化填充配置
{
    public class Sample01
    {
        public static string Run()
        {
            var result = EPPlusHelper.GetFillDefaultConfig("序号	工号	姓名	性别");
            Console.WriteLine(result);
            return result;
        }

    }
}
