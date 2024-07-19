namespace EPPlusExtensions
{
    public class EPPlusReport
    {
        /// <summary>
        /// 行折叠的deth(Level)在DataRow中什么列表示.默认:Level
        /// </summary>
        public string RowLevelColumnName { get; set; } = "Level";

        #region 行折叠默认配置

        /// <summary>
        /// 行是否合并显示
        /// </summary>
        public bool Collapsed { get; set; } = true;

        /// <summary>
        /// 合并/展开 行 的折叠符号是否在右边
        /// </summary>
        public bool OutLineSummaryBelow { get; set; } = false;

        #endregion

        #region 列 折叠默认配置(暂不支持该功能,主要是还没遇到列合并的导出需求)

        ///// <summary>
        ///// 列 是否合并显示
        ///// </summary>
        //public bool RowCollaspsed { get; set; } = true;
        ///// <summary>
        ///// 合并/展开 列 的折叠符号是否在下面
        ///// </summary>
        //public bool OutLineSummaryRight { get; set; } = false;

        #endregion
    }
}
