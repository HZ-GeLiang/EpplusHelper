namespace EPPlusExtensions
{
    /// <summary>
    /// 数据模板配置信息
    /// </summary>
    public class ExcelDataConfigInfo
    {
        /// <summary>
        /// 从1开始
        /// </summary>
        public int WorkSheetIndex { get; set; }

        /// <summary>
        /// Index 和 Name 填写一个就可以了
        /// </summary>
        public string WorkSheetName { get; set; }

        /// <summary>
        /// 标题行(对于合并单元格,写起始单元格的信息)
        /// </summary>
        public int TitleLine { get; set; }

        /// <summary>
        /// 标题列(对于合并单元格,写起始单元格的信息)
        /// </summary>
        public int TitleColumn { get; set; }
    }
}