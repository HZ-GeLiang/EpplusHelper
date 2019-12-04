using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp._01填充数据
{
    class Sample00
    {
        internal static DataTable GetDataTable_Head()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Title");
            DataRow dr = dt.NewRow();
            dr["Title"] = "2018第一学期考试";
            dt.Rows.Add(dr);
            return dt;

        }
        internal static DataTable GetDataTable_Body()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Chinese");
            dt.Columns.Add("Math");
            dt.Columns.Add("English");
            for (int i = 0; i < 5; i++)
            {
                DataRow dr = dt.NewRow();
                dr["Name"] = $"张三{i + 1}";
                dr["Chinese"] = 60;
                dr["Math"] = 60.5;
                dr["English"] = 61;
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
