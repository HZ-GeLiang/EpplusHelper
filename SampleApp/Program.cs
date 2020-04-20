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
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using SampleApp._01填充数据;
using SampleApp.MethodExtension;

namespace SampleApp
{

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

            SampleApp._03读取excel内容.Sample01.Run();
            //SampleApp._03读取excel内容.Sample02.Run();
            //SampleApp._03读取excel内容.Sample03.Run();
            //SampleApp._03读取excel内容.Sample04.Run();
            //SampleApp._03读取excel内容.Sample04_2.Run();
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


 
}
