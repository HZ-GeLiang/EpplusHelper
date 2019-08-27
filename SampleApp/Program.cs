using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;


namespace SampleApp
{


    class Program
    {

        static void Main(string[] args)
        {
            new Sample01_1().Run();
        }
    }

    public class 内部推荐岗位信息
    {
        public string 序号 { get; set; }
        public string 岗位名称 { get; set; }
        public string 岗位描述 { get; set; }
        public string 技能要求 { get; set; }
        public string 简历推荐人 { get; set; }
    }


}
