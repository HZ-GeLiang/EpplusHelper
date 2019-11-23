﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions;
using EPPlusExtensions.Attributes;
using OfficeOpenXml;

namespace SampleApp
{
    /// <summary>
    /// 读取Excel的内容
    /// </summary>
    class Sample02_1_2
    {
        public void Run()
        {
            string filePath = @"模版\Sample02_1.xlsx";
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                try
                {
                    ExcelWorksheet ws = EPPlusHelper.GetExcelWorksheet(excelPackage, "Sheet1");
                    var args = EPPlusHelper.GetExcelListArgsDefault<ysbm>(ws, 2);
                    args.ScanLine = ScanLine.SingleLine;

                    var propModel = new ysbm();


                    var source = propModel.部门.CreateKVSource();

                    {
                        //添加方式1
                        var dataSource = propModel.部门.CreateKVSourceData();
                        dataSource.Add("事业1部", 1);
                        dataSource.Add("事业2部", 2);
                        dataSource.Add("事业3部", 3);
                        source.AddRange(dataSource);
                    }
                    {
                        //添加方式2
                        #region dt
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

                        foreach (DataRow item in dt.Rows)
                        {
                            //确保类型是对的
                            var key = SafeRow(item, "Name", propModel.部门.GetKeyType());
                            var value = SafeRow(item, "Id", propModel.部门.GetValueType());
                            source.TryAdd(key, value);
                        }
                    }

                    {
                        //添加方式3
                        #region dt
                        var dt = new DataTable();
                        dt.Columns.Add("Id");
                        dt.Columns.Add("Name");

                        var dr3 = dt.NewRow();
                        dr3["Id"] = 6;
                        dr3["Name"] = "事业6部";

                        dt.Rows.Add(dr3);
                        #endregion

                        source.AddRange(GetSource_部门(propModel, dt).Data);
                    }
                    args.KVSource.Add(nameof(propModel.部门), source);

                    var list = EPPlusHelper.GetList<ysbm>(args);

                    ObjectDumper.Write(list);
                    Console.WriteLine("读取完毕");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            Console.WriteLine("按任意键结束程序!");
            Console.ReadKey();

        }

        public static object SafeRow(DataRow row, string name, Type type)
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

        private static KvSource<string, long> GetSource_部门(ysbm propModel, DataTable dt)
        {
            var prop = propModel.部门;
            KvSource<string, long> kvsource = prop.CreateKVSource();
            foreach (DataRow item in dt.Rows)
            {
                //var key = SafeRow(item, "Name", prop.KeyType);
                //var value = SafeRow(item, "Id", prop.ValueType);

                var key = SafeRow(item, "Name", propModel.部门.GetKeyType());
                var value = SafeRow(item, "Id", propModel.部门.GetValueType());
                kvsource.TryAdd(key, value);
            }
            return kvsource;
        }


        internal class ysbm
        {
            public string 序号 { get; set; }
            //[KVSet("部门")] // 属性'部门'值:'事业1部'未在'部门'集合中出现
            //[KVSet("部门", "部门在数据库中未找到")] //部门在数据库中未找到
            [KVSet("部门", "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            //[KVSet("部门", false, "'{0}'在数据库中未找到", "部门")]//'事业1部'在数据库中未找到
            public KV<string, long> 部门 { get; set; }
            public string 部门负责人 { get; set; }
            public string 部门负责人确认签字 { get; set; }
        }
    }
}
