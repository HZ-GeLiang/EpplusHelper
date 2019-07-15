using OfficeOpenXml;

namespace EPPlusExtensions
{
    /// <summary>
    /// 单元格信息
    /// </summary>
    public class ExcelCellInfo
    {
        public ExcelWorksheet WorkSheet { get; set; }
        public ExcelAddress ExcelAddress { get; set; }
        public object Value { get; set; }
    }

    public class ExcelCellInfoValue
    {
        /// <summary>
        /// 符合c#命名规范的名字
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 有没有重命名
        /// </summary>
        public bool IsRename { get; set; }
        /// <summary>
        /// 新的名字
        /// </summary>
        public string NameNew { get; set; }
        /// <summary>
        /// 每个列名的Col位置
        /// </summary>
        public int ExcelColNameIndex { get; set; }
        /// <summary>
        /// excel原本的列名
        /// </summary>
        public string ExcelColName { get; set; }
    }
}
