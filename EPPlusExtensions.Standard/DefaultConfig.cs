namespace EPPlusExtensions
{
    public class DefaultConfig
    {
        /// <summary>
        /// 工作簿名字
        /// </summary>
        public string WorkSheetName { get; set; }

        /// <summary>
        /// 代码片段-创建DataTable
        /// </summary>
        public string CrateDataTableSnippe { get; set; }

        /// <summary>
        /// 代码片段-创建类的
        /// </summary>
        public string CrateClassSnippe { get; set; }

        /// <summary>
        /// Class的属性列表
        /// </summary>
        public List<ExcelCellInfoValue> ClassPropertyList { get; set; }
    }
}
