using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;

namespace SampleApp
{
    /// <summary>
    /// 自动初始化填充配置
    /// </summary>
    class Sample04_1
    {
        public void Run()
        {
            var result = EPPlusHelper.GetFillDefaultConfig("序号	工号	姓名	性别");
            Console.WriteLine(result);
        }

    }
}
