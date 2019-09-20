using System;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    /// <summary>
    /// 合并单元格
    /// </summary>
    public struct ExcelCellRange
    {
        private static void Init(string r1c1,
            out string Range,
            out ExcelCellPoint Start,
            out ExcelCellPoint End,
            out int IntervalRow,
            out int IntervalCol,
            out bool IsMerge)
        {
            Range = r1c1;
            var cellPoints = r1c1.Split(':');

            switch (cellPoints.Length)
            {
                case 1:
                    Start = new ExcelCellPoint(cellPoints[0].Trim());
                    End = default(ExcelCellPoint);
                    IntervalCol = 0;
                    IntervalRow = 0;
                    IsMerge = false;
                    break;
                case 2:
                    Start = new ExcelCellPoint(cellPoints[0].Trim());
                    End = new ExcelCellPoint(cellPoints[1].Trim());
                    IntervalCol = End.Col - Start.Col;
                    IntervalRow = End.Row - Start.Row;
                    IsMerge = true;
                    break;
                default:
                    throw new Exception("程序的配置有问题");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="r1c1">地址</param>
        public ExcelCellRange(string r1c1)
        {
            Init(r1c1,
                out string Range,
                out ExcelCellPoint Start,
                out ExcelCellPoint End,
                out int IntervalRow,
                out int IntervalCol,
                out bool IsMerge);

            this.Range = Range;
            this.Start = Start;
            this.End = End;
            this.IntervalRow = IntervalRow;
            this.IntervalCol = IntervalCol;
            this.IsMerge = IsMerge;
        }

        public ExcelCellRange(string r1c1, ExcelWorksheet ws)
        {
            var ecp = new ExcelCellPoint(r1c1);
            var ea = new ExcelAddress(ws.MergedCells[ecp.Row, ecp.Col]);

            Init(ea.Address,
                out string Range,
                out ExcelCellPoint Start,
                out ExcelCellPoint End,
                out int IntervalRow,
                out int IntervalCol,
                out bool IsMerge);
            this.Range = Range;
            this.Start = Start;
            this.End = End;
            this.IntervalRow = IntervalRow;
            this.IntervalCol = IntervalCol;
            this.IsMerge = IsMerge;
        }

        /// <summary>
        /// 范围也就是R1C1的地址(保存的是配置时的字符串.在程序中用来当作key使用)
        /// </summary>
        public string Range { get; private set; }

        /// <summary>
        /// 开始Point
        /// </summary>
        public ExcelCellPoint Start { get; private set; }

        /// <summary>
        /// 结束Point
        /// </summary>
        public ExcelCellPoint End { get; private set; }

        /// <summary>
        /// 间距行是多少
        /// </summary>
        public int IntervalRow { get; private set; }

        /// <summary>
        /// 间距列是多少
        /// </summary>
        public int IntervalCol { get; private set; }

        /// <summary>
        /// 是否是合并单元格
        /// </summary>
        public bool IsMerge { get; private set; }
    }
}
