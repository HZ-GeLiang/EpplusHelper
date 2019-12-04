using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp._02填充图片
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
            dt.Columns.Add("Evaluate");

            DataRow dr = dt.NewRow();
            dr["Name"] = "张三";
            dr["Chinese"] = 60;
            dr["Math"] = 60.5;
            dr["English"] = 61;
            dr["Evaluate"] = CaptchaGen.ImageFactory.GenerateImage("合", 50, 100, 13, 0);
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr["Name"] = "李四";
            dr["Chinese"] = 70;
            dr["Math"] = 80.5;
            dr["English"] = 91;
            dr["Evaluate"] = CaptchaGen.ImageFactory.GenerateImage("优", 50, 100, 13, 0);
            dt.Rows.Add(dr);

            return dt;

        }
    }
}
