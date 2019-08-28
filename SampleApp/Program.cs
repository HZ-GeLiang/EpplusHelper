using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using SampleApp.MethodExtension;


namespace SampleApp
{


    class Program
    {


        static void Main(string[] args1)
        {
            //string filePath = @"C:\Users\child\Desktop\a.xlsx";
            //using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //using (ExcelPackage excelPackage = new ExcelPackage(fs))
            //{
            //    ExcelWorksheet worksheet = EPPlusHelper.GetExcelWorksheet(excelPackage, 1, "导出测试");
            //    worksheet.InsertRow(3, 2);
            //    excelPackage.SaveAs(ms);
            //    ms.Position = 0;
            //    ms.Save(@"C:\Users\child\Desktop\a1.xlsx");
            //}
             
            new Sample01_1().Run();
        }
        public enum RecruitCandidate_ResumeFrom
        {
            内部推荐 = 1,
            智联招聘 = 2,
            前程无忧 = 3,
            Boss直聘 = 4,
            RPO = 5,
            拉勾 = 6,
            猎聘 = 7,
            猎头 = 8,
            人才库回访 = 9,
            重新聘用 = 10,
            内部提拔 = 11,
            内部转岗 = 12,
            邮箱 = 13,
            实习僧 = 14,
            校园招聘 = 15,
        }
        public enum RecruitInterviewResult_RecruitInterviewResult
        {
            通过 = 1,
            不通过 = 2
        }
        public enum RecruitRecommendTalent_RecruitDecision
        {
            邀约 = 1,
            不邀约 = 2,

        }
        public class EM14面试评价表
        {
            public string 姓名 { get; set; }
            public string 性别 { get; set; }
            public string 手机号 { get; set; }
            public string 邮箱 { get; set; }
            public string 应聘岗位 { get; set; }
            public RecruitCandidate_ResumeFrom 简历来源 { get; set; }
            public string 简历来源详细 { get; set; }

            [Required]
            public string 招聘负责人 { get; set; }
            [Required]
            public string 筛选人 { get; set; }


            public RecruitRecommendTalent_RecruitDecision 决定 { get; set; }
            public string HR面试官 { get; set; }
            public DateTime? HR面试日期 { get; set; }
            public string HR面试时间 { get; set; }
            public string 专业面试官 { get; set; }
            public DateTime? 专业面日期 { get; set; }
            public string 专业面时间 { get; set; }
            public string 管理面试官 { get; set; }
            public DateTime? 管理面试官日期 { get; set; }
            public string 管理面试官时间 { get; set; }
            public RecruitInterviewResult_RecruitInterviewResult? HR面试结果 { get; set; }
            public RecruitInterviewResult_RecruitInterviewResult? 专业面试官面试结果 { get; set; }
            public RecruitInterviewResult_RecruitInterviewResult? 管理面试官面试结果 { get; set; }
            /// <summary>
            /// 含 - 中 +  范围
            /// </summary>
            public string 能力等级 { get; set; }
            public string HR面试评价 { get; set; }
            public string 专业面试官面试评价 { get; set; }
            public string 管理面试官面试评价 { get; set; }
            public string 学历 { get; set; }
            public string 公司经历 { get; set; }
            public string 其他 { get; set; }
        }
    }



}
