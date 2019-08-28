using System;
using System.Collections.Generic;
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
            Col = R1C1Formulas(RegexHelper.GetFirstStringByReg(r1C1, "[A-Za-z]+"));//K -> 11
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
            R1C1 = R1C1FormulasReverse(col) + row;
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
            Col = R1C1Formulas(RegexHelper.GetFirstStringByReg(r1C1, "[A-Za-z]+"));//K -> 11
            R1C1 = r1C1;
        }

        /// <summary>
        /// 譬如: A->1 . 在excel的选项->属性->公式  下有个 R1C1引用样式
        /// </summary>
        /// <param name="col">只能是字母</param>
        /// <returns></returns>
        public static int R1C1Formulas(string col)
        {
            col = col.ToUpper();
            Dictionary<string, int> r1C1 = new Dictionary<string, int>
            {
                {"A", 1},{"B", 2},{"C", 3},{"D", 4},{"E", 5},{"F", 6},
                {"G", 7},{"H", 8},{"I", 9},{"J", 10},{"K", 11},{"L", 12},
                {"M", 13},{"N", 14},{"O", 15},{"P", 16},{"Q", 17},{"R", 18},
                {"S", 19},{"T", 20},{"U", 21},{"V", 22},{"W", 23},{"X", 24},
                {"Y", 25},{"Z", 26},
            };
            int colLength = col.Length;
            if (colLength == 1)
            {
                return r1C1[col];
            }
            int sum = 0;
            for (int i = 0; i < colLength; i++)
            {
                char c = col[i];
                int num = r1C1[c + ""];
                sum += (int)(num * Math.Pow(26, colLength - i - 1));
            }
            return sum;
        }

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
            Dictionary<int, char> r1C1 = new Dictionary<int, char>
            {
                {1,'A'},{2,'B'},{3,'C'},{4,'D'},{5,'E'},{6,'F'},
                {7,'G'},{8,'H'},{9,'I'},{10,'J'},{11,'K'},{12,'L'},
                {13,'M'},{14,'N'},{15,'O'},{16,'P'},{17,'Q'},{18,'R'},
                {19,'S'},{20,'T'},{21,'U'},{22,'V'},{23,'W'},{24,'X'},
                {25,'Y'},{26,'Z'},{0,'A'}
            };
            if (num <= 26) //这个if属于优化,若删掉.也没有关系
            {
                return r1C1[num].ToString();
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
                    charList.Add(mod != 0 ? r1C1[mod] : r1C1[26]);
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
