using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EpplusExtensions;
using EpplusExtensions.Attributes;
using OfficeOpenXml;


namespace SampleApp
{
    class Program
    {
      

        class Model1
        {
            public string 序号 { get; set; }
            //public string 姓名 { get; set; }
            [ExcelColumnIndex(3)]
            [DisplayExcelColumnName("请假次数")]
            public string 请假次数1 { get; set; }
            [ExcelColumnIndex(4)]
            [DisplayExcelColumnName("请假次数")]
            public string 请假次数2 { get; set; }
        }
        static void Main(string[] args)
        {
            new Sample02_5().Run();


            string filePath = $@"C:\Users\child\Desktop\Sample02_7.xlsx";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EpplusHelper.GetExcelWorksheet(excelPackage, 1);
                List<Model1> list;
                try
                {
                    list = EpplusHelper.GetList<Model1>(ws, 2);
                    ObjectDumper.Write(list);
                }
                catch (Exception e)
                {
                    throw e;
                }
                Console.WriteLine("读取完毕");
            }
          
            //OpenDirectoryHelp.OpenFilePath(System.IO.Path.Combine(OpenDirectoryHelp.GetSaveFilePath(), @"Debug\模版\"));
            //new Sample01_1().Run();
            //new Sample01_2().Run();
            //new Sample01_1_2().Run();
            //new Sample02_1().Run();
            ////new Sample02_2().Run();
            ////new Sample02_3().Run();
            //new Sample02_4().Run();
            //new Sample02_5().Run();
            ////new Sample03_1().Run();
            ////new Sample03_2().Run();
            //new Sample04_1().Run();
            //new Sample04_2().Run();
            //new Sample04_3().Run(); 
        }
    }
}
