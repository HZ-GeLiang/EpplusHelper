using EpplusExtensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp
{
    /// <summary>
    /// 填充数据与数据源同步
    /// </summary>
    class Sample04_2
    {
        public void Run()
        {

            string str = $@"序号	工号	姓名	性别	入司时间	转正时间	离职时间	离职类型	归属公司	部门	二级部门	三级部门	组别	岗位	岗位大类	行政职级	离职约谈记录
          ";
            var result = EpplusHelper.GetFillConfig(str);
            Console.WriteLine(result);
        }

    }
}
