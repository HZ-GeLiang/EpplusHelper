using EPPlusExtensions.Helpers;
using System.Text;

namespace EPPlusExtensions
{
    /// <summary>
    /// 普通单元格
    /// </summary>
    public struct ExcelCellPoint
    {
        public int Row;
        public int Col;
        public string ColStr;

        /// <summary>
        /// 譬如A2等
        /// </summary>
        public string R1C1;

        ///// <summary>
        /////
        ///// </summary>
        ///// <param name="row">从1开始的整数</param>
        ///// <param name="col">只能是字母</param>
        ///// <param name="r1C1">譬如A2 等</param>
        //public ExcelCellPoint(int row, string col, string r1C1)
        //{
        //    Row = row;
        //    Col = R1C1Formulas(col);
        //    R1C1 = r1C1;
        //}
        public ExcelCellPoint(string r1C1)
        {
            //K3 = row:3, col:11
            r1C1 = r1C1.Split(':')[0].Trim(); //防止传入 "A1:B3" 这种的配置格式的
            Row = Convert.ToInt32(RegexHelper.GetLastNumber(r1C1));//3
            ColStr = RegexHelper.GetFirstStringByReg(r1C1, "[A-Za-z]+");
            Col = ExcelCellPoint.R1C1Formulas(ColStr);//K -> 11
            R1C1 = r1C1;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="row">从1开始的整数</param>
        /// <param name="col">从1开始的整数</param>
        public ExcelCellPoint(int row, int col)
        {
            Row = row;
            Col = col;
            ColStr = ExcelCellPoint.R1C1FormulasReverse(col);
            R1C1 = ExcelCellPoint.R1C1FormulasReverse(col) + row;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="excelAddress"></param>
        public ExcelCellPoint(OfficeOpenXml.ExcelAddress excelAddress)
        {
            //ExcelCellPoint(excelAddress.Address);
            var r1C1 = excelAddress.Address;
            r1C1 = r1C1.Split(':')[0].Trim(); //防止传入 "A1:B3" 这种的配置格式的
            Row = Convert.ToInt32(RegexHelper.GetLastNumber(r1C1));//3
            ColStr = RegexHelper.GetFirstStringByReg(r1C1, "[A-Za-z]+");
            Col = ExcelCellPoint.R1C1Formulas(ColStr);//K -> 11
            R1C1 = r1C1;
        }

        private static Dictionary<char, int> r1C1_Dict = new Dictionary<char, int>(26 * 2)
        {
            {'A', 01},{'B', 02},{'C', 03},{'D', 04},{'E', 05},{'F', 06},{'G', 07},
            {'H', 08},{'I', 09},{'J', 10},{'K', 11},{'L', 12},{'M', 13},{'N', 14},
            {'O', 15},{'P', 16},{'Q', 17},{'R', 18},{'S', 19},{'T', 20},{'U', 21},
            {'V', 22},{'W', 23},{'X', 24},{'Y', 25},{'Z', 26},
            {'a', 01},{'b', 02},{'c', 03},{'d', 04},{'e', 05},{'f', 06},{'g', 07},
            {'h', 08},{'i', 09},{'j', 10},{'k', 11},{'l', 12},{'m', 13},{'n', 14},
            {'o', 15},{'p', 16},{'q', 17},{'r', 18},{'s', 19},{'t', 20},{'u', 21},
            {'v', 22},{'w', 23},{'x', 24},{'y', 25},{'z', 26},
        };

        /// <summary>
        /// 譬如: A->1 . 在excel的选项->属性->公式  下有个 R1C1引用样式
        /// </summary>
        /// <param name="col">只能是字母</param>
        /// <returns></returns>
        public static int R1C1Formulas(string col)
        {
            if (string.IsNullOrEmpty(col))
            {
                throw new ArgumentNullException(nameof(col));
            }

            int sum = 0;
            for (int i = 0; i < col.Length; i++)
            {
                char c = col[i];
                int num = ExcelCellPoint.r1C1_Dict[c];
                sum += (int)(num * Math.Pow(26, col.Length - i - 1));
            }
            return sum;
        }

        private static Dictionary<int, char> r1C1_reverse_dict = new Dictionary<int, char>(27)
        {
            {01,'A'},{02,'B'},{03,'C'},{04,'D'},{05,'E'},{06,'F'},{07,'G'},
            {08,'H'},{09,'I'},{10,'J'},{11,'K'},{12,'L'},{13,'M'},{14,'N'},
            {15,'O'},{16,'P'},{17,'Q'},{18,'R'},{19,'S'},{20,'T'},{21,'U'},
            {22,'V'},{23,'W'},{24,'X'},{25,'Y'},{26,'Z'},{00,'A'}
        };

        /// <summary>
        /// 譬如1->A
        /// </summary>
        /// <param name="num">excel的第几列</param>
        /// <returns></returns>
        public static string R1C1FormulasReverse(int num)
        {
            if (num <= 0)
            {
                throw new Exception("parameter 'col' can not less zero");
            }

            if (num <= 26) //这个if属于优化,若删掉.也没有关系
            {
                return ExcelCellPoint.r1C1_reverse_dict[num].ToString();
            }
            List<char> charList = new List<char>();
            int cimi = -1; //次幂数

            while (true)
            {
                cimi++;
                var num2 = (long)Math.Pow(26, cimi); //while的终止条件与计算条件
                if (num >= num2)
                {
                    int mod = num / (int)num2 % 26;//余数
                    charList.Add(mod != 0 ? ExcelCellPoint.r1C1_reverse_dict[mod] : ExcelCellPoint.r1C1_reverse_dict[26]);
                    num -= (int)num2;
                }
                else
                {
                    break;
                }
            }

            StringBuilder sb = new StringBuilder();
            for (int i = charList.Count - 1; i >= 0; i--)
            {
                sb.Append(charList[i]);
            }
            return sb.ToString();
        }
    }
}