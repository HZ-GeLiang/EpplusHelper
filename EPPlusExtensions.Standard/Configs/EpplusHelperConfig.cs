namespace EPPlusExtensions
{
    public class EpplusHelperConfig
    {
        public static List<string> KeysTypeOfDecimal = new List<string>
        {
            "金额", "钱", "数额",
            "money", "Money", "MONEY",
            "amount", "Amount", "AMOUNT",
        };


        public static List<string> KeysTypeOfDateTime => new List<string>
        {
            "时间", "日期", "date", "Date", "DATE", "time", "Time", "TIME",
            "今天", "昨天", "明天", "前天",
            "day", "Day", "DAY",
            "tomorrow","Tomorrow","TOMORROW",
        };


        public static List<string> KeysTypeOfString = new List<string>
        {
            "序号", "编号", "id", "Id", "ID", "number", "Number", "NUMBER", "No",
            "身份证", "银行卡", "卡号", "手机", "座机",
            "mobile", "Mobile", "MOBILE", "tel", "Tel", "TEL",
        };

    }
}