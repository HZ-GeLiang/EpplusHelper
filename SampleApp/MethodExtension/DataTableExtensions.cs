using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SampleApp.MethodExtension
{
    internal static class DataTableExtensions
    {
        /// <summary>
        /// 常用字符串
        /// </summary>
        public class CharConst
        {
            /// <summary>
            /// HT (horizontal tab) 水平制表符 ascii 9
            /// </summary>
            // ReSharper disable once InconsistentNaming
            public const char HT = '\t'; //"	";

            /// <summary>
            /// VT (vertical tab) 垂直制表符   ascii 11
            /// </summary>
            // ReSharper disable once InconsistentNaming
            public const char VT = '\v';

            /// <summary>
            /// 空格符号  ascii 160  = ((char)160).ToString()
            /// </summary>
            public const string Space160 = " ";
        }

        /// <summary>
        /// 常用字符串
        /// </summary>
        public class StringConst
        {
            /// <summary>
            /// HT (horizontal tab) 水平制表符 ascii 9
            /// </summary>
            // ReSharper disable once InconsistentNaming
            public const string HT = "\t";//"	";

            /// <summary>
            /// VT (vertical tab) 垂直制表符   ascii 11
            /// </summary>
            // ReSharper disable once InconsistentNaming
            public const string VT = "\v";

            /// <summary>
            /// 空格符号  ascii 160  = ((char)160).ToString()
            /// </summary>
            public const string Space160 = " ";
        }

        /// <summary>
        /// DataTable 转成 List&lt;T&gt;若DataTable为空,不返回null(可放心使用linq)
        /// </summary>
        /// <typeparam name="T">泛型类型</typeparam>
        /// <returns>List &lt;T&gt;</returns>
        public static List<T> ToList<T>(this DataTable dt)
        {
            if (dt is null || dt.Rows.Count <= 0)
            {
                return Enumerable.Empty<T>().ToList();
            }

            List<T> list = Enumerable.Empty<T>().ToList();   //创建泛型集合对象
            //遍历数据行，将行数据存入 实体对象中，并添加到 泛型集合中list

            Type t = typeof(T);//1先获得泛型的类型
            T model = (T)Activator.CreateInstance(t); //2根据类型创建该类型的对象
            PropertyInfo[] properties = t.GetProperties();  //3根据类型 获得 该类型的 所有属性定义

            foreach (DataRow row in dt.Rows)
            {
                model = (T)Activator.CreateInstance(t); //每次虚幻创建一个对象
                foreach (PropertyInfo p in properties)  //4遍历属性数组
                {
                    string colName = p.Name;  //4.1获得属性名，作为列名
                    object colValue = row[colName]; //4.2根据列名 获得当前循环行对应列的值
                    if (colValue != DBNull.Value)
                    {
                        var isEnum = p.PropertyType.IsEnum;
                        if (isEnum)
                        {
                            p.SetValue(model, Enum.ToObject(p.PropertyType, colValue));
                        }
                        else
                        {
                            try
                            {
                                p.SetValue(model, colValue);  //4.3将 列值 赋给 model对象的p属性
                            }
                            catch (System.ArgumentException e)
                            {
#if DEBUG
                                Console.WriteLine(e.Message);
#endif
                                //数据库类型和model类型不一致
                                if (p.PropertyType == typeof(string))
                                {
                                    p.SetValue(model, colValue.ToString());
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                }
                list.Add(model); //5将装好 了行数据的 实体对象 添加到 泛型集合中
            }
            return list;
        }

        public static string ToText(this DataTable dt, bool needTitle = true)
        {
            if (dt is null || dt.Rows.Count <= 0)
            {
                return string.Empty;
            }

            var sbTxt = new StringBuilder();

            if (needTitle && dt.Columns.Count > 0)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    sbTxt.Append(column.ColumnName).Append(StringConst.HT);
                }
                sbTxt.RemoveLastChar(CharConst.HT);
                sbTxt.Append(Environment.NewLine);
            }

            //遍历数据行，将行数据存入 实体对象中，并添加到 泛型集合中list
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn column in dt.Columns)
                {
                    sbTxt.Append(row[column.ColumnName]).Append(StringConst.HT);
                }
                sbTxt.RemoveLastChar(CharConst.HT);
                sbTxt.Append(Environment.NewLine);
            }

            var txt = sbTxt.RemoveLastChar(Environment.NewLine).ToString();
            return txt;
        }
    }
}