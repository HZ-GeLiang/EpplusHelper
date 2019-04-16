using System;
using EPPlusExtensions;

namespace SampleApp.Core
{
    /// <summary>
    /// 自动初始化填充配置
    /// </summary>
    class Sample04_1
    {
        public void Run()
        {
            string str = $@"序号	工号	姓名	性别	入司时间	转正时间	离职时间	离职类型	归属公司	部门	二级部门	三级部门	组别	岗位	岗位大类	行政职级	离职约谈记录";
            var result = EPPlusHelper.GetFillDefaultConfig(str);
            Console.WriteLine(result);
        }

    }
}
