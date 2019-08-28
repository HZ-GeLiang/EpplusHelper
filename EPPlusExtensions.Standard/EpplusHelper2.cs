using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using EPPlusExtensions.Attributes;
using EPPlusExtensions.Exceptions;
using EPPlusExtensions.Helper;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    public partial class EPPlusHelper
    {
        #region 对单元格样式进行 Get Set

        ///// <summary>
        /////  获取Cell样式
        ///// </summary>
        ///// <param name="cell"></param>
        ///// <returns></returns>
        //public static EPPlusCellStyle GetCellStyle(ExcelRange cell)
        //{
        //    EPPlusCellStyle cellStyle = new EPPlusCellStyle();
        //    cellStyle.HorizontalAlignment = cell.Style.HorizontalAlignment;
        //    cellStyle.VerticalAlignment = cell.Style.VerticalAlignment;
        //    cellStyle.WrapText = cell.Style.WrapText;
        //    cellStyle.FontBold = cell.Style.Font.Bold;
        //    cellStyle.FontColor = string.IsNullOrEmpty(cell.Style.Font.Color.Rgb)
        //        ? Color.Black
        //        : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Font.Color.Rgb);
        //    cellStyle.FontName = cell.Style.Font.Name;
        //    cellStyle.FontSize = cell.Style.Font.Size;
        //    cellStyle.BackgroundColor = string.IsNullOrEmpty(cell.Style.Fill.BackgroundColor.Rgb)
        //        ? Color.Black
        //        : System.Drawing.ColorTranslator.FromHtml("#" + cell.Style.Fill.BackgroundColor.Rgb);
        //    cellStyle.ShrinkToFit = cell.Style.ShrinkToFit;
        //    return cellStyle;
        //}

        ///// <summary>
        ///// 设置Cell样式
        ///// </summary>
        ///// <param name="cell"></param>
        ///// <param name="style"></param>
        //public static void SetCellStyle(ExcelRange cell, EPPlusCellStyle style)
        //{
        //    cell.Style.HorizontalAlignment = style.HorizontalAlignment;
        //    cell.Style.VerticalAlignment = style.VerticalAlignment;
        //    cell.Style.WrapText = style.WrapText;
        //    cell.Style.Font.Bold = style.FontBold;
        //    cell.Style.Font.Color.SetColor(style.FontColor);
        //    if (!string.IsNullOrEmpty(style.FontName))
        //    {
        //        cell.Style.Font.Name = style.FontName;
        //    }
        //    cell.Style.Font.Size = style.FontSize;
        //    cell.Style.Fill.PatternType = style.PatternType;
        //    cell.Style.Fill.BackgroundColor.SetColor(style.BackgroundColor);
        //    cell.Style.ShrinkToFit = style.ShrinkToFit;
        //}

        #endregion

        #region 一些默认的sql语句,SqlServer 下使用

        /// <summary>
        /// 获得树形表结构的最深的层级数的Sql语句
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="rootItemWhere">root(根)数据的where条件,即根据表名获得root(根)数据的条件是什么</param>
        public static string GetTreeTableMaxLevelSql(string tblName, string rootItemWhere, string idFiledName = "Id", string parentIdName = "ParentId")
        {
            string sql = $@"
with cte as( 
    SELECT {idFiledName} ,  1 as level FROM {tblName} WHERE {rootItemWhere}
    UNION ALL
    SELECT {tblName}.{idFiledName}, cte.level+1 as level from cte, {tblName}  where cte.{idFiledName} = {tblName}.{parentIdName} 
)
SELECT ISNULL(MAX(cte.level),0) FROM  cte";
            return sql;
        }

        /// <summary>
        /// 原本的树形表结构是没有Level字段的,通过该方法可以生成level字段
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="rootItemWhere"></param>
        /// <param name="nameFieldName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="otherFiledName"></param>
        /// <returns></returns>
        public static string GetTreeTableIncludeLevelFieldSql(string tblName, string rootItemWhere, string nameFieldName = "Name", string idFiledName = "Id", string parentIdName = "ParentId", params string[] otherFiledName)
        {
            string comma = " ,";
            string dot = ".";
            StringBuilder sb1 = new StringBuilder(); //定位成员的字段
            sb1.Append(idFiledName).Append(comma)
                .Append(nameFieldName).Append(comma)
                .Append(parentIdName);
            StringBuilder sb2 = new StringBuilder(); //递归成员的字段
            sb2.Append(tblName).Append(dot).Append(idFiledName).Append(comma)
                .Append(tblName).Append(dot).Append(nameFieldName).Append(comma)
                .Append(tblName).Append(dot).Append(parentIdName);

            if (otherFiledName != null && otherFiledName.Length > 0)
            {
                foreach (var item in otherFiledName)
                {
                    sb1.Append(item).Append(comma);
                    sb2.Append(tblName).Append(dot).Append(item).Append(comma);
                }
                sb1.RemoveLastChar(comma.Length);
                sb2.RemoveLastChar(comma.Length);
            }

            string sql = $@"
with cte as( 
    SELECT {sb1} , 1 as Level FROM {tblName} WHERE {rootItemWhere}
    UNION ALL
    SELECT {sb2} , cte.Level+1 as Level from cte, {tblName}  
        where cte.{idFiledName} = {tblName}.{parentIdName} 
)
SELECT {sb1} , Level FROM  cte
ORDER BY cte.Level";
            return sql;
        }

        /// <summary>
        ///  根据 id, Name , parentId 3个字段生成额外字段Depth 和 用于报表排序的Sort字段
        /// </summary>
        /// <param name="tblName"></param>
        /// <param name="rootItemWhere"></param>
        /// <param name="nameFieldName"></param>
        /// <param name="idFiledName"></param>
        /// <param name="parentIdName"></param>
        /// <param name="eachSortFieldLength">每个Depth的长度,默认2. </param>
        /// <param name="reportSortFileTotallength">报表排序字段的总长度,默认为12如果真的要设置,level * Max(Len(主键))</param>
        /// <param name="rearChat">报表排序字段 / 每个Depth字段 小于 指定长度时填充的字符是什么</param>
        /// <param name="otherFiledName"></param>
        /// <returns></returns>
        public static string GetTreeTableReportSql(string tblName, string rootItemWhere, string nameFieldName = "Name", string idFiledName = "Id", string parentIdName = "ParentId", int eachSortFieldLength = 2, int reportSortFileTotallength = 12, char rearChat = ' ', params string[] otherFiledName)
        {
            //该方法基本与GetTreeTableIncludeLevelFieldSql()一样
            string comma = " ,";
            string dot = ".";
            StringBuilder sb1 = new StringBuilder(); //定位成员的字段
            sb1.Append(idFiledName).Append(comma)
                .Append(nameFieldName).Append(comma)
                .Append(parentIdName);
            StringBuilder sb2 = new StringBuilder(); //递归成员的字段
            sb2.Append(tblName).Append(dot).Append(idFiledName).Append(comma)
                .Append(tblName).Append(dot).Append(nameFieldName).Append(comma)
                .Append(tblName).Append(dot).Append(parentIdName);

            string char1 = Enumerable.Repeat(rearChat.ToString(), eachSortFieldLength).Aggregate((current, next) => next + current);
            string char2 = Enumerable.Repeat(rearChat.ToString(), reportSortFileTotallength).Aggregate((current, next) => next + current);

            if (otherFiledName != null && otherFiledName.Length > 0)
            {
                foreach (var item in otherFiledName)
                {
                    sb1.Append(item).Append(comma);
                    sb2.Append(tblName).Append(dot).Append(item).Append(comma);
                }
                sb1.RemoveLastChar(comma.Length);
                sb2.RemoveLastChar(comma.Length);
            }

            string sql = $@"
with cte as( 
    SELECT {sb1} , 1 as Level , CAST( LEFT(LTRIM({idFiledName})+'{char1}',{eachSortFieldLength}) AS VARCHAR(10)) AS 'Depth'
    FROM {tblName} WHERE {rootItemWhere}
    UNION ALL
    SELECT {sb2} , cte.Level+1 as Level , CAST(LTRIM(cte.Depth) + LEFT(LTRIM({tblName}.{idFiledName}) +'{char1}',{eachSortFieldLength})AS VARCHAR(10)) AS 'Depth' 
    FROM cte, {tblName} 
    where cte.{idFiledName} = {tblName}.{parentIdName} 
)
SELECT {sb1} , Level,LEFT(LTRIM(cte.Depth)+'{char2}',{reportSortFileTotallength})  AS 'sort'  FROM cte
ORDER BY sort ,cte.Level";
            return sql;

        }

        #endregion 


        /// <summary>
        /// 
        /// </summary>
        /// <param name="action"></param>
        /// <returns></returns>
        public static string GetListErrorMsg(Action action)
        {
            try
            {
                action.Invoke();
                return null;
            }
            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("程序报错:");
                if (e.Message != null && e.Message.Length > 0)
                {
                    sb.AppendLine($@"Message:{e.Message}");
                }
                if (e.InnerException != null && e.InnerException.Message != null && e.InnerException.Message.Length > 0)
                {
                    sb.AppendLine($@"InnerExceptionMessage:{e.InnerException.Message}");
                }

                return sb.ToString();
            }
        }
    }
}
