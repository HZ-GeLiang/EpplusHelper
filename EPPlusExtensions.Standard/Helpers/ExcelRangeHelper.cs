using OfficeOpenXml;

namespace EPPlusExtensions
{
    public sealed class ExcelRangeHelper
    {
        public static string GetCellText(ExcelRange cell, bool when_TextProperty_NullReferenceException_Use_ValueProperty = true)
        {
            //if (cell.Merge) throw new Exception("没遇到过这个情况的");
            //return cell.Text; //这个没有科学计数法  注:Text是Excel显示的值,Value是实际值.
            try
            {
                //if (cell.Formula?.Length > 0)//cell 是公式
                //{
                //}
                return cell.Text;//有的单元格通过cell.Text取值会发生异常,但cell.Value却是有值的

                //例如，如果你在单元格中输入日期"2024-04-14"并将其格式化为日期格式，
                //Excel将会在"Text"中显示"2024-04-14"，但在"Value"中存储对应的序列号（如45396）。
                //详见示例07

                /*
                我没遇到过这个场景, 这个代码先保留

                if (cell.IsRichText)
                {
                    //https://www.cnblogs.com/studyever/archive/2012/08/29/2661850.html
                    return cell.RichText.Text;
                }
                */
            }
            catch (NullReferenceException)
            {
                if (when_TextProperty_NullReferenceException_Use_ValueProperty)
                {
                    return Convert.ToString(cell.Value);
                }
                throw;
            }
        }

        /// <summary>
        /// 设置单元格的的值
        /// </summary>
        /// <param name="cell">目前针对的场景是非合并单元格, 如果是合并单元格, 没测试过</param>
        /// <param name="cellValue"></param>
        public static void SetWorksheetCellValue(ExcelRange cell, string cellValue)
        {
            cell.Value = cellValue;
            if (string.IsNullOrWhiteSpace(cellValue) == false && object.Equals(cellValue, cell.Value) == false) // 有值,但没有填充上去
            {
                if (cell.IsRichText)
                {
                    cell.RichText.Text = cellValue;
                }
            }
        }

    }
}