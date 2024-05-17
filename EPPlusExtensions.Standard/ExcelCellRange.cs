using System;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    /// <summary>
    /// 合并单元格
    /// </summary>
    public struct ExcelCellRange
    {

        /// <summary>
        ///
        /// </summary>
        /// <param name="r1c1">地址</param>
        public ExcelCellRange(string r1c1) : this(r1c1, null)
        {

        }

        public ExcelCellRange(string r1c1, ExcelWorksheet ws)
        {
            string _r1c1;
            if (ws != null)
            {
                var ecp = new ExcelCellPoint(r1c1);
                if (!EPPlusHelper.IsMergeCell(ws, ecp.Row, ecp.Col, out var mergeCellAddress))
                {
                    throw new Exception($@"r1c1:{r1c1}不是合并单元格");
                }
                var ea = new ExcelAddress(mergeCellAddress);
                _r1c1 = ea.Address;
            }
            else
            {
                _r1c1 = r1c1;
            }

            this.Range = _r1c1;
            var cellPoints = _r1c1.Split(':');

            if (cellPoints.Length == 1)
            {
                this.Start = new ExcelCellPoint(cellPoints[0].Trim());
                this.End = default(ExcelCellPoint);
                this.IntervalCol = 0;
                this.IntervalRow = 0;
                this.IsMerge = false;
            }
            else if (cellPoints.Length == 2)
            {
                this.Start = new ExcelCellPoint(cellPoints[0].Trim());
                this.End = new ExcelCellPoint(cellPoints[1].Trim());
                this.IntervalCol = End.Col - Start.Col;
                this.IntervalRow = End.Row - Start.Row;
                this.IsMerge = true;
            }
            else
            {
                throw new Exception("程序的配置有问题");
            }
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
