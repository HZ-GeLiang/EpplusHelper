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
            string tempPath = $@"模版\Sample04_2.xlsx";
            EpplusHelper.FillExcelDefaultConfig(tempPath, Path.GetDirectoryName(Path.GetFullPath(tempPath)));
        }
    }
}
