using System;

namespace EpplusExtensions
{
    /// <summary>
    /// 合并单元格
    /// </summary>
    public struct ExcelCellRange
    {
        public ExcelCellRange(string range)
        {
            Range = range;
            var cellPoints = range.Split(':');

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
        /// 范围(保存的是配置时的字符串.在程序中用来当作key使用)
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
