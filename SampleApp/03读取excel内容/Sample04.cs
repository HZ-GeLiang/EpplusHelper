using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace SampleApp._03读取excel内容
{
    public class Sample04
    {
        public static List<ExcelModel> Run()
        {
            string filePath = @"模版\03读取excel内容\Sample04.xlsx";
            var wsName = "合并行读取";
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fs))
            {
                var ws = EPPlusHelper.GetExcelWorksheet(excelPackage, wsName);
                var args = EPPlusHelper.GetExcelListArgsDefault<ExcelModel>(ws, 2);

                var source = args.Model.部门.CreateKVSource();
                AddSourceWay1_AddRange_CreateDataSource(args.Model, source);
                AddSourceWay2_TryAdd_CreateDataTable(args.Model, source);
                AddSourceWay3_AddRange_ByFunction(args.Model, source);

                args.Model.部门.KVSource = source;
                args.Model.部门评分.KVSource = GetSource_部门评分(args.Model);
                var list = EPPlusHelper.GetList(args).ToList();
                ObjectDumper.Write(list);
                Console.WriteLine("读取完毕");
                return list;
            }
        }

        private static KvSource<long, string> GetSource_部门评分(ExcelModel propModel)
        {
            var source = propModel.部门评分.CreateKVSource();
            source.Add(1, "非常不满意");
            source.Add(2, "不满意");
            source.Add(3, "一般");
            source.Add(4, "满意");
            source.Add(5, "非常满意");
            return source;
        }


        /// <summary>
        /// 内部TryAdd,自己封装一个方法, 个人推荐用这个.代码改的少
        /// </summary>
        /// <param name="propModel"></param>
        /// <param name="source"></param>
        private static void AddSourceWay3_AddRange_ByFunction(ExcelModel propModel, KvSource<string, long> source)
        {
            #region dt

            var dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");

            var dr = dt.NewRow();
            dr["Id"] = 6;
            dr["Name"] = "事业6部";

            dt.Rows.Add(dr);

            #endregion

            //source.AddRange(GetSource_部门(propModel, dt).Data);

            var prop = propModel.部门;
            var keyType = prop.GetKeyType();
            var valueType = prop.GetValueType();

            var kvsource = prop.CreateKVSource();
            foreach (DataRow item in dt.Rows)
            {
                var key = SafeRow(item, "Name", keyType);
                var value = SafeRow(item, "Id", valueType);
                kvsource.TryAdd(key, value);
            }

            source.AddRange(kvsource.Data);
        }

        /// <summary>
        ///  TryAdd 自己创建 DataTable
        /// </summary>
        /// <param name="propModel"></param>
        /// <param name="source"></param>
        private static void AddSourceWay2_TryAdd_CreateDataTable(ExcelModel propModel, KvSource<string, long> source)
        {
            #region CreateDataTable

            var dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            var dr1 = dt.NewRow();
            var dr2 = dt.NewRow();
            dr1["Id"] = 4;
            dr1["Name"] = "事业4部";
            dr2["Id"] = 5;
            dr2["Name"] = "事业5部";
            dt.Rows.Add(dr1);
            dt.Rows.Add(dr2);

            #endregion

            var prop = propModel.部门;
            var keyType = prop.GetKeyType();
            var valueType = prop.GetValueType();

            foreach (DataRow item in dt.Rows)
            {
                //确保类型是对的
                var key = SafeRow(item, "Name", keyType);
                var value = SafeRow(item, "Id", valueType);
                source.TryAdd(key, value);
            }
        }

        /// <summary>
        /// KV添加方式1-AddRange() 自己创建 DataSource
        /// </summary>
        /// <param name="propModel"></param>
        /// <param name="source"></param>
        private static void AddSourceWay1_AddRange_CreateDataSource(ExcelModel propModel, KvSource<string, long> source)
        {
            #region CreatDataSource

            var dataSource = propModel.部门.CreateKVSourceData();

            //dataSource.Add("事业1部", 1);
            //dataSource.Add("事业2部", 2);
            //dataSource.Add("事业3部", 3);

            #region dt

            var dt = new DataTable();
            dt.Columns.Add("Id");
            dt.Columns.Add("Name");
            var dr1 = dt.NewRow();
            var dr2 = dt.NewRow();
            var dr3 = dt.NewRow();
            dr1["Id"] = 1;
            dr1["Name"] = "事业1部";
            dr2["Id"] = 2;
            dr2["Name"] = "事业2部";
            dr3["Id"] = 3;
            dr3["Name"] = "事业3部";
            dt.Rows.Add(dr1);
            dt.Rows.Add(dr2);
            dt.Rows.Add(dr3);

            #endregion

            //var prop = propModel.部门;
            //var keyType = prop.GetKeyType();
            //var valueType = prop.GetValueType();

            foreach (DataRow item in dt.Rows)
            {
                //var k = SafeRow(item, "Name", keyType);
                //var v = SafeRow(item, "Id", valueType);
                //dataSource.Add((string)k, (long)v);

                dataSource.Add(item["Name"].ToString(), Convert.ToInt64(item["Id"]));
            }

            #endregion

            source.AddRange(dataSource);
        }

        static object SafeRow(DataRow row, string name, Type type)
        {
            object o = row[name];
            if (o == DBNull.Value || o == null)
            {
                return null;
            }
            if (o is string && ((string)o).Length == 0)
            {
                if (type == typeof(string))
                {
                    return o;
                }
                else if (type == typeof(bool))
                {
                    return default(bool);
                }
                else if (type == typeof(decimal))
                {
                    return default(decimal);
                }
                else if (type == typeof(double))
                {
                    return default(double);
                }
                else if (type == typeof(float))
                {
                    return default(float);
                }
                else if (type == typeof(DateTime))
                {
                    return default(DateTime);
                }
                else if (type == typeof(Int64))
                {
                    return default(Int64);
                }
                else if (type == typeof(Int32))
                {
                    return default(Int32);
                }
                else if (type == typeof(Int16))
                {
                    return default(Int16);
                }
                else
                {
                    throw new Exception("请完善改方法");
                }
            }
            return Convert.ChangeType(o, type);

        }

        static KvSource<string, long> GetSource_部门(ExcelModel propModel, DataTable dt)
        {
            var prop = propModel.部门;
            var keyType = prop.GetKeyType();
            var valueType = prop.GetValueType();

            var kvsource = prop.CreateKVSource();
            foreach (DataRow item in dt.Rows)
            {
                var key = SafeRow(item, "Name", keyType);
                var value = SafeRow(item, "Id", valueType);
                kvsource.TryAdd(key, value);
            }
            return kvsource;
        }


        public class ExcelModel
        {

            public string 序号 { get; set; }

            [KVSet("'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            //[KVSet("部门", false, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long> 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }

            [KVSet]
            public KV<long, string> 部门评分 { get; set; }

            public override bool Equals(object obj)
            {
                if (obj == null || !obj.GetType().Equals(this.GetType()))
                {
                    return false;
                }

                ExcelModel y = (ExcelModel)obj;

                return this.序号 == y.序号 &&
                       Helper.GetEquals_KV(this.部门, y.部门) &&
                       this.部门负责人 == y.部门负责人 &&
                       this.部门负责人确认签字 == y.部门负责人确认签字 &&
                       Helper.GetEquals_KV(this.部门评分, y.部门评分);
            }

            //重写Equals方法必须重写GetHashCode方法，否则发生警告
            public override int GetHashCode()
            {
                return this.序号.GetHashCode() +
                       Helper.GetHashCode_KV(this.部门) +
                       Helper.GetHashCode_KV(this.部门评分);
            }

        }
    }
}
