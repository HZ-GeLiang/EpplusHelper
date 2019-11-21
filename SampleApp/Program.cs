﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using SampleApp.MethodExtension;


namespace SampleApp
{
    //Func<float, Func<int, float>> happyWater = new Func<float, int, float>((price, number) => number * price).Currying();
    //Func<float, int, float> happyWater2 = new Func<float, int, float>((price, number) => number * price);

    class Program
    {
        static void Main(string[] args1)
        {
            new Sample02_1_2().Run();
        }
    }
     

    public static class CurryingExtensions
    {
        //https://mp.weixin.qq.com/s?__biz=MzAxMTMxMDQ3Mw==&mid=2660105542&idx=1&sn=9519dc358cde59e1c6d27773007d5699&chksm=803a59a0b74dd0b6c8a54d3b0967c5bbf7a7c8e92bc3867cd0d099dc08bb8f6e6a4be4c23881&scene=0&xtrack=1&key=0b6f00fa5c3dca5d8719d70beb5e2fecd35b4d8cfb2f28b7c55737d2cb9e2d2b677bb0d6ee198169e333ad0d16dd0c208befe018725150cd96494049cfd155a423dc435f191349d522125d06b3e0fe60&ascene=1&uin=MTgyMTkyNzMwMg%3D%3D&devicetype=Windows+10&version=62060834&lang=zh_CN&pass_ticket=J7b3DfTgb3w9fp7EBZI7udUSW58lTVIRztEd0OMKb6fh%2B0bx100d9R77pES6VeYd

        public static Func<T1, Func<T2, TOutput>> Currying<T1, T2, TOutput>(this Func<T1, T2, TOutput> f) => x => y => f(x, y);
        public static Func<T1, Func<T2, Func<T3, TOutput>>> Currying<T1, T2, T3, TOutput>(this Func<T1, T2, T3, TOutput> f) => x => y => z => f(x, y, z);
    }
}
