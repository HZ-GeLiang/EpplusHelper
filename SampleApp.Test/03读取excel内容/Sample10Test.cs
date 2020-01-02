using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApp._03读取excel内容;
using EPPlusExtensions.Attributes;

namespace SampleApp.Test._03读取excel内容
{
    [TestClass]
    public class Sample10Test
    {
        [TestMethod]
        public void TestMethod1()
        {
            var dt = new DataTable();
            dt.Columns.Add("序号");
            dt.Columns.Add("部门");
            dt.Columns.Add("部门负责人");
            dt.Columns.Add("部门负责人确认签字");

            var dr = dt.NewRow();
            dr["序号"] = 1; dr["部门"] = "事业1部"; dr["部门负责人"] = "赵六"; dr["部门负责人确认签字"] = "娃娃";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["序号"] = 2; dr["部门"] = "事业2部"; dr["部门负责人"] = "赵六"; dr["部门负责人确认签字"] = "菲菲";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            //dr["序号"] = 3; dr["部门"] = "事业3部"; dr["部门负责人"] = "王五"; dr["部门负责人确认签字"] = "佩琪";
            //dt.Rows.Add(dr);
            //dr = dt.NewRow();
            //dr["序号"] = 4; dr["部门"] = "事业4部"; dr["部门负责人"] = "jam"; dr["部门负责人确认签字"] = "jam";
            //dt.Rows.Add(dr);
            //dr = dt.NewRow();
            //dr["序号"] = 6; dr["部门"] = "事业6部"; dr["部门负责人"] = "jack"; dr["部门负责人确认签字"] = "jack";
            //dt.Rows.Add(dr);

            var dt2 = Sample10.Run();

            if (dt2 == null)
            {
                Assert.Fail("dt2返回了Null");
            }
            if (dt.Rows.Count != dt2.Rows.Count)
            {
                Assert.Fail("DataTable的记录行数不一样");
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Assert.AreEqual(dt.Rows[i][j].ToString(), dt2.Rows[i][j].ToString());
                }
            }

        }
    }
}
