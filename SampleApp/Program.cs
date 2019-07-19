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
            string FilePath = @"C:\Users\child\Desktop\有问题.xlsx";
            var WsName = "2019下半年需求";
            using (MemoryStream ms = new MemoryStream())
            using (FileStream fs = new System.IO.FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, WsName);
                var list = EPPlusHelper.GetList<EM_ZPJHDR>(ws, 2);
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
            }
        }
    }

    public enum UserInfo_Company
    {
        畅唐网络 = 1,
        昊汉网络 = 2,
        星罗恒渡 = 3,
        逍遥网络 = 5,
        崇汉科技 = 6,
        美则美矣 = 7,
        星罗集团 = 8,
        星罗控股 = 9,
        昊汉代理商 = 10, // 0x0000000A
        易秦网络 = 11, // 0x0000000B
        星罗普渡 = 13, // 0x0000000D
        金唐网络 = 14, // 0x0000000E
        印度天秦 = 15, // 0x0000000F
    }
    public class EM_ZPJHDR
    {
        public string 序号 { get; set; }
        public DateTime 需求提交日期 { get; set; }
        public UserInfo_Company 隶属公司 { get; set; }
        public string 一级部门 { get; set; }
        public string 二级部门 { get; set; }
        public string 三级部门 { get; set; }
        public string 岗位名称 { get; set; }
        public string 岗位大类 { get; set; }
        public RecruitRequireApply_RequirementType 需求性质 { get; set; }
        public RecruitRequireApply_RequirementType 岗位性质 { get; set; }

        [ExcelColumnIndex(11)]
        [DisplayExcelColumnName("需求P级")]
        public PMLlevel 需求P级1 { get; set; }
        [ExcelColumnIndex(12)]
        [DisplayExcelColumnName("需求P级")]
        public PMLlevel 需求P级2 { get; set; }
        public string 需求到岗月份 { get; set; }
        public string 招聘负责人 { get; set; }
        public 入职表Id? 姓名 { get; set; }
        public DateTime 到岗时间 { get; set; }
        public string 岗位小类 { get; set; }
        public string 职级 { get; set; }
        [DisplayExcelColumnName("定薪（转正）")]
        public string 定薪转正 { get; set; }
        public string 辅导员 { get; set; }
        public string 专业面试官 { get; set; }
        public string 定薪主管 { get; set; }
        [DisplayExcelColumnName("补员对象（选填）")]
        public string 补员对象选填 { get; set; }
        public string 备注 { get; set; }
        public string 需求提出人 { get; set; }


    }
    public enum SexNoLimit
    {
        男 = 1,
        女 = 2,
        不限 = 3,
    }
    public enum RecruitRequireApply_RequirementType
    {
        新增需求 = 1,
        离职补员 = 2,
        调岗补员 = 3,

    }
    /// <summary>
    /// 性别
    /// </summary>
    public enum Sex
    {
        /// <summary>
        /// 男性
        /// </summary>
        男 = 1,
        /// <summary>
        /// 女性
        /// </summary>
        女 = 2,
    }


    public enum PMLlevel
    {
        无 = 25,
        P1 = 1,
        P2 = 2,
        P3 = 3,
        P4 = 4,
        P5 = 5,
        P6 = 6,
        P7 = 7,
        P8 = 8,
        P9 = 9,
        P10 = 10,
        P11 = 11,
        P12 = 12,
        M1 = 13,
        M2 = 14,
        M3 = 15,
        M4 = 16,
        M5 = 17,
        M6 = 18,
        M7 = 19,
    }

    public class 内部推荐岗位信息
    {
        public string 序号 { get; set; }
        public string 岗位名称 { get; set; }
        public string 岗位描述 { get; set; }
        public string 技能要求 { get; set; }
        public string 简历推荐人 { get; set; }
    }

    public enum 入职表Id : long
    {
        //SELECT tueac.Name,'=',tueac.Id,',' FROM ctoa.tblUserEntryApplyChild AS tueac WHERE tueac.Status =1
        余力来 = 202027926333,
        张瑶瑶 = 202027926555,
        王雪松 = 202028155183,
        宋吉 = 202028466093,
        卢燕萍 = 202029882853,
        张书桂 = 202032184137,
        夏磊 = 202100001313,
        刘珂 = 202100853674,
        朱璐瑶 = 202104414891,
        廖伟麒 = 202105261057,
        胡瑛骏 = 202105261279,
        陆书琛 = 202108285357,
        李敬欢 = 202109050040,
        金李广 = 202109537349,
        王韵慧 = 202110384741,
        吴爱玲 = 202112139349,
        李欣蓉 = 202112152871,
        毕琳崧 = 202112153048,
        周豪邦 = 202113215752,
        姚超群 = 202113241617,
        朱杰 = 202113753967,
        叶盼盼 = 202114356661,
        宋岩 = 202114937077,
        李佳 = 202115157286,
        杜恺珺 = 202115206061,
        毕加跃 = 202115206212,
        邢兴中 = 202115206483,
        李洪博 = 202115850780,
        张海勇 = 202115850953,
        陈俐蓉 = 202115851175,
        芮箕环 = 202116640362,
        丁英英 = 202116958487,
        马志云 = 202117464144,
        任萍 = 202117464369,
        李杰 = 202117948970,
        易盼 = 202117949126,
        林峰 = 202118173649,
        黄玉蓉 = 202118282418,
        王星 = 202118826118,
        孙毓 = 202119459977,
        廖剑斌 = 202119622024,
        俞瑜 = 202119648530,
        陈佳源 = 202119884555,
        何旭杭 = 202119884933,
        付智超 = 202120002288,
        彭杰 = 202120664979,
        张钰 = 202121008460,
        李佳瑶 = 202121008677,
        袁明俊 = 202121240763,
        诸晓锋 = 202121405457,
        吴江铭 = 202121405617,
        王浩 = 202121405835,
        陈叶龙 = 202121406012,
        武继路 = 202122337147,
        严海燕 = 202122337365,
        钟志杰 = 202122563439,
        蔡庆秀 = 202122563657,
        程映华 = 202122983566,
        徐佳洋 = 202123417677,
        王银婷 = 202123417887,
        马琪凯 = 202123418025,
        吴二峰 = 202124022356,
        涂运 = 202124773378,
        缪明远 = 202124773569,
        邢力翔 = 202124773734,
        冯士顺 = 202124773942,
        王博 = 202125423759,
        郑宇峰 = 202126050226,
        胡宇斐 = 202126050466,
        唐枫 = 202126050649,
        胡鑫 = 202126050874,
        李芳 = 202126051024,
        盛凤琦 = 202126051296,
        何娇 = 202126439216,
        陈龙 = 202126890344,
        徐春林 = 202127382574,
        刘宏 = 202127382756,
        夏敏杰 = 202127383179,
        贺茂纯 = 202127741165,
        樊静 = 202127741356,
        周海海 = 202128185979,
        郭光伟 = 202128410369,
        王洁 = 202128410594,
        段君保 = 202128410745,
        彭娟娟 = 202128605227,
        王火东 = 202128936977,
        陈果 = 202129078024,
        楼天一 = 202129078264,
        崔凤璇 = 202129692923,
        方力 = 202129839636,
        陈杭彬 = 202130078175,
        景悦 = 202130243862,
        程旭 = 202130244457,
        孙聪聪 = 202130244674,
        沈盼盼 = 202130244849,
        孔泽锋 = 202130542082,
        任晓阳 = 202130737648,
        刘金亮 = 202130737822,
        刘彬彬 = 202130738025,
        陈开樑 = 202131222282,
        崔启 = 202131222443,
        林斯盛 = 202131979310,
        陈东 = 202132706794,
        李俊林 = 202132706954,
        徐楠 = 202132707172,
        归涛 = 202133239797,
        施竹羽 = 202133239989,
        富浩 = 202133602258,
        许斌 = 202133602450,
        姚婉 = 202133602628,
        邵晶 = 202133602867,
        阮家恩 = 202133819625,
        吕心标 = 202134047149,
        刘少华 = 202134673236,
        鲁永吉 = 202134673428,
        韩进平 = 202134673638,
        杨龙龙 = 202134853610,
        余小雨 = 202134853840,
        胡春亮 = 202135108639,
        林小捷 = 202135290539,
        申静伟 = 202135331459,
        叶安 = 202136158936,
        齐文灿 = 202136174682,
        赵旭辉 = 202136449042,
        贾林昌 = 202136764198,
        吴江红 = 202136764381,
        孙张欢 = 202136958617,
        范哲宁 = 202136958857,
        韩凯丰 = 202137186327,
        孟久翔 = 202137186598,
        郑巨隆 = 202137371277,
        毛鸿鸽 = 202137628042,
        张明欢 = 202137628225,
        卢敏强 = 202138464732,
        汪昌月 = 202138464989,
        韩天枫 = 202139063282,
        陈懿 = 202139063465,
        张永健 = 202139063615,
        毛伟伟 = 202139755992,
        匡贤旺 = 202139756174,
        杜俊震 = 202139756310,
        叶雯彦 = 202139756582,
        巩宇翔 = 202140339389,
        於刚 = 202140339525,
        李倩 = 202140519171,
        杨晨光 = 202140519353,
        施青楠 = 202140519593,
        韩家辉 = 202140519776,
        苏杨 = 202140774492,
        章胜男 = 202140774612,
        黑后 = 202141531755,
        李五七 = 202141536659,
        李五八 = 202141537492,
        二号 = 202141556730,
    }

}
