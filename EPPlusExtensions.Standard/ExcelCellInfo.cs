using OfficeOpenXml;

namespace EPPlusExtensions
{
    /// <summary>
    /// 单元格信息
    /// </summary>
    public class ExcelCellInfo
    {
        public ExcelWorksheet WorkSheet { get; set; }
        public ExcelCellPoint ExcelCellPoint { get; set; }
        public object Value { get; set; }
    }
}
