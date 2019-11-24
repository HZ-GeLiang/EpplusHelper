using EPPlusExtensions.Attributes;
using EPPlusExtensions.Exceptions;
using EPPlusExtensions.Helper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlusExtensions
{
    public partial class EPPlusHelper
    {
        #region GetList<T>

        public static GetExcelListArgs<T> GetExcelListArgsDefault<T>(ExcelWorksheet ws, int rowIndex) where T : class
        {
            return new GetExcelListArgs<T>()
            {
                ws = ws,
                RowIndex_Data = rowIndex,
                EveryCellPrefix = "",
                EveryCellReplaceList = null,
                RowIndex_DataName = rowIndex - 1,
                UseEveryCellReplace = true,
                HavingFilter = null,
                WhereFilter = null,
                ReadCellValueOption = ReadCellValueOption.Trim,
                POCO_Property_AutoRename_WhenRepeat = false,
                POCO_Property_AutoRenameFirtName_WhenRepeat = true,
                ScanLine = ScanLine.MergeLine,
                MatchingModelEqualsCheck = true,
                GetList_NeedAllException = false,
                GetList_ErrorMessage_OnlyShowColomn = false,
                DataColStart = 1,
                DataColEnd = EPPlusConfig.MaxCol07,
                KVSource = new Dictionary<string, object>(),
            };
        }

        /// <summary>
        /// 只能是最普通的excel.(每个单元格都是未合并的,第一行是列名,数据从第一列开始填充的那种.)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowIndex">数据起始行(不含列名),从1开始</param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex) where T : class, new()
        {
            var args = EPPlusHelper.GetExcelListArgsDefault<T>(ws, rowIndex);
            return EPPlusHelper.GetList<T>(args);
        }

        /// <summary>
        /// 只能是最普通的excel.(第一行是必须是列名,数据填充列起始必须是A2单元格,且每个单元格都是未合并的)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ws"></param>
        /// <param name="rowIndex">数据起始行(不含列名),从1开始</param> 
        /// <param name="everyCellReplaceOldValue"></param>
        /// <param name="everyCellReplaceNewValue"></param>
        /// <returns></returns>
        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellReplaceOldValue, string everyCellReplaceNewValue) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowIndex);
            if (everyCellReplaceOldValue != null && everyCellReplaceNewValue != null)
            {
                args.EveryCellReplaceList = new Dictionary<string, string> { { everyCellReplaceOldValue, everyCellReplaceNewValue } };
            }
            return GetList<T>(args);
        }

        public static List<T> GetList<T>(ExcelWorksheet ws, int rowIndex, string everyCellPrefix, Dictionary<string, string> everyCellReplace) where T : class, new()
        {
            var args = GetExcelListArgsDefault<T>(ws, rowIndex);
            args.EveryCellPrefix = everyCellPrefix;
            args.EveryCellReplaceList = everyCellReplace;
            return GetList<T>(args);
        }

        public static List<T> GetList<T>(GetExcelListArgs<T> args) where T : class, new()
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.RowIndex_Data;
            if (rowIndex <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowIndex}'错误,值应该大于0");
            }

            int rowIndex_DataName = args.RowIndex_DataName;
            if (rowIndex_DataName <= 0)
            {
                throw new ArgumentException($@"数据起始行的标题行值'{rowIndex_DataName}'错误,值应该大于0");
            }

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, args.DataColStart, args.DataColEnd, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
            if (colNameList.Count == 0)
            {
                throw new Exception("未读取到单元格标题");
            }
            var dictExcelAddressCol = colNameList.ToDictionary(item => item.ExcelAddress, item => new ExcelCellPoint(item.ExcelAddress).Col);

            Type type = typeof(T);

            #region 获得字典
            var dictModelPropNameExistsExcelColumn = new Dictionary<string, bool>();//Model属性在Excel列中存在, key: ModelPropName
            var dictModelPropNameToExcelColumnName = new Dictionary<string, string>();//Model属性名字对应的excel的标题列名字
            var dictExcelColumnIndexToModelPropName_Temp = new Dictionary<int, string>();//Excel的列标题和Model属性名字的映射
            var dictExcelColumnIndexToModelPropName_All = new Dictionary<int, string>();//Excel列对应的Model属性名字(所有excel列)

            foreach (var props in type.GetProperties())
            {
                dictModelPropNameExistsExcelColumn.Add(props.Name, false);
                dictModelPropNameToExcelColumnName.Add(props.Name, null);

                var propAttr_DisplayExcelColumnName = ReflectionHelper.GetAttributeForProperty<DisplayExcelColumnNameAttribute>(type, props.Name);
                if (propAttr_DisplayExcelColumnName.Length > 0)
                {
                    dictModelPropNameToExcelColumnName[props.Name] = ((DisplayExcelColumnNameAttribute)propAttr_DisplayExcelColumnName[0]).Name;
                }
                var propAttr_ExcelColumnIndex = ReflectionHelper.GetAttributeForProperty<ExcelColumnIndexAttribute>(type, props.Name);
                if (propAttr_ExcelColumnIndex.Length > 0)
                {
                    dictExcelColumnIndexToModelPropName_Temp.Add(((ExcelColumnIndexAttribute)propAttr_ExcelColumnIndex[0]).Index, props.Name);
                }

            }

            foreach (var item in colNameList)
            {
                //var excelColumnIndex = new ExcelCellRange(item.ExcelAddress.ToString()).Start.Col;
                var excelColumnIndex = dictExcelAddressCol[item.ExcelAddress];
                dictExcelColumnIndexToModelPropName_All.Add(excelColumnIndex, null);
                string propName = item.Value.ToString();
                PropertyInfo pInfo = type.GetProperty(propName);
                if (pInfo == null)
                {
                    if (dictExcelColumnIndexToModelPropName_Temp.ContainsKey(excelColumnIndex))
                    {
                        var propNameTemp = dictExcelColumnIndexToModelPropName_Temp[excelColumnIndex];
                        //不做属性的 DisplayExcelColumnName = 当前属性的验证 (因为还没想到这个属性是一定要验证的情况)
                        //if (dictModelPropNameToExcelColumnName.ContainsKey(propNameTemp) && dictModelPropNameToExcelColumnName[propNameTemp] == propName)
                        //{
                        //   pInfoTemp = type.GetProperty(propName);
                        //}

                        var pInfoTemp = type.GetProperty(propNameTemp);
                        if (pInfoTemp != null)
                        {
                            propName = propNameTemp;
                            pInfo = pInfoTemp;
                        }
                    }
                }

                if (pInfo != null)
                {
                    dictModelPropNameExistsExcelColumn[propName] = true;
                    dictExcelColumnIndexToModelPropName_All[excelColumnIndex] = propName;
                }
            }
            #endregion

            #region 验证 MatchingModel.eq //args.MatchingModel

            var _matchingModelSuccess = false;  //提供的 Matching 参数[这里写死了MatchingModel.eq] 和算出来MatchingModel 有没有交集(默认没有)

            var dictExcelColumnIndexToExcelColName = colNameList.ToDictionary(item => new ExcelCellPoint(item.ExcelAddress).Col, item => item.Value.ToString());
            var _matchingModel = GetMatchingModel(dictExcelColumnIndexToExcelColName, dictExcelColumnIndexToModelPropName_All,
                dictModelPropNameExistsExcelColumn, out List<string> modelPropNotExistsExcelColumn, out List<string> excelColumnIsNotModelProp);
            var _matchingModelValues = Enum.GetValues(typeof(MatchingModel));
            foreach (MatchingModel matchingModelValue in _matchingModelValues)
            {
                if ((MatchingModel.eq & matchingModelValue) == matchingModelValue &&
                    (_matchingModel & matchingModelValue) == matchingModelValue) //提供的 Matching 参数[这里写死了MatchingModel.eq] 和算出来Matching 有重叠
                {
                    _matchingModelSuccess = true;
                    break;
                }
            }
            if (!_matchingModelSuccess)
            {
                #region 获得 listMatchingModelException

                var dictMatchingModelException = new Dictionary<MatchingModel, MatchingModelException>() { };
                //var colNameToCellInfo = colNameList.ToDictionary(item => item.Value.ToString(), item => item);//当excel列有重复,Model没用Attribute,导致Modle 与Excel不匹配, 此时会报错, Demo 在Sample02_7 , 把Model 的Attribute可以复现

                var colNameToCellInfo = new Dictionary<string, List<ExcelCellInfo>>();

                foreach (var colName in colNameList)
                {
                    var colIndex = dictExcelAddressCol[colName.ExcelAddress];
                    var modelpropName = dictExcelColumnIndexToModelPropName_All[colIndex];
                    if (modelpropName != null)
                    {
                        if (!colNameToCellInfo.ContainsKey(modelpropName))
                        {
                            colNameToCellInfo.Add(modelpropName, new List<ExcelCellInfo> { });
                        }
                        colNameToCellInfo[modelpropName].Add(colName);
                    }
                    else
                    {
                        var excelColVaue = colName.Value.ToString();
                        if (!colNameToCellInfo.ContainsKey(excelColVaue))
                        {
                            colNameToCellInfo.Add(excelColVaue, new List<ExcelCellInfo> { });
                            colNameToCellInfo[excelColVaue].Add(colName);
                        }
                        else
                        {
                            ////暂时不考虑多次提供了不存在的列的情况, 即: 不存在的列只能多一共一次,否则报错
                            //throw new Exception($@"当前Excel多次提供了,根据值:{excelColVaue},在Model中找不对应属性,当前列是:{new ExcelCellRange(colName.ExcelAddress.ToString()).Start.R1C1}");
                            if (!colNameToCellInfo.ContainsKey(excelColVaue))
                            {
                                colNameToCellInfo.Add(excelColVaue, new List<ExcelCellInfo> { });
                            }
                            colNameToCellInfo[excelColVaue].Add(colName);
                        }
                    }

                }

                foreach (MatchingModel matchingModelValue in _matchingModelValues)
                {
                    if (matchingModelValue == MatchingModel.eq)
                    {
                        #region excel的哪些列与Model不相等

                        if ((_matchingModel & MatchingModel.eq) != MatchingModel.eq) continue;
                        if (dictMatchingModelException.ContainsKey(matchingModelValue)) continue;//如果已经添加过了

                        if (excelColumnIsNotModelProp.Count == 0 && modelPropNotExistsExcelColumn.Count == 0)
                        {
                            dictMatchingModelException.Add(MatchingModel.eq, new MatchingModelException()
                            {
                                MatchingModel = MatchingModel.eq,
                                ListExcelCellInfoAndModelType = null
                            });
                        }
                        else
                        {
                            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
                            foreach (var colName in excelColumnIsNotModelProp)
                            {
                                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType()
                                {
                                    ExcelCellInfoList = colNameToCellInfo[colName],
                                    ModelType = type
                                });
                            }

                            dictMatchingModelException.Add(MatchingModel.eq, new MatchingModelException()
                            {
                                MatchingModel = MatchingModel.eq,
                                ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType
                            });
                        }

                        #endregion
                    }
                    else if (matchingModelValue == MatchingModel.gt)
                    {
                        if ((_matchingModel & MatchingModel.gt) != MatchingModel.gt) continue;
                        if (dictMatchingModelException.ContainsKey(MatchingModel.gt)) continue;
                        dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, ws));
                    }
                    else if (matchingModelValue == MatchingModel.lt)
                    {
                        if ((_matchingModel & MatchingModel.lt) != MatchingModel.lt) continue;
                        if (dictMatchingModelException.ContainsKey(MatchingModel.lt)) continue;
                        dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, ws));
                    }
                    else if (matchingModelValue == MatchingModel.neq)
                    {
                        if ((_matchingModel & MatchingModel.neq) != MatchingModel.neq) continue;
                        //neq 会调用 gt+ lt ,所以要排除,即 _matchingModel的值 不能是带neq的标志枚举的值
                        if ((_matchingModel & MatchingModel.gt) == MatchingModel.gt) continue;
                        if ((_matchingModel & MatchingModel.lt) == MatchingModel.lt) continue;

                        #region excel的哪些列与Model不相等

                        //excel的哪些列 不在Model中定义+ model中定义了,但是excel列中却没有 

                        if (!dictMatchingModelException.ContainsKey(MatchingModel.gt))
                        {
                            dictMatchingModelException.Add(MatchingModel.gt, GetMatchingModelExceptionCase_gt(modelPropNotExistsExcelColumn, type, colNameToCellInfo, ws));
                        }
                        if (!dictMatchingModelException.ContainsKey(MatchingModel.lt))
                        {
                            dictMatchingModelException.Add(MatchingModel.lt, GetMatchingModelExceptionCase_lt(excelColumnIsNotModelProp, type, colNameToCellInfo, ws));
                        }

                        #endregion
#if DEBUG
                        if ((_matchingModel & MatchingModel.eq) == MatchingModel.eq)
                        {
                            throw new Exception("断言:这里应该是不会进来的,debug下调试看看,进来是什么情况");
                        }
#endif

                    }
                    else
                    {
                        throw new Exception("不支持的MatchingMode值");
                    }
                }

                #endregion

                StringBuilder sb = new StringBuilder();

                //foreach (var matchingModelException in listMatchingModelException)
                foreach (var matchingModelException in dictMatchingModelException.Values)
                {
                    var errMsg = DealMatchingModelException(matchingModelException);
                    sb.Append(errMsg);
                }
                if (sb.Length > 0)
                {
                    throw new MatchingModelException(sb.ToString());
                }

                throw new Exception("验证未通过,程序有bug");

            }

            #endregion

            //var dictColName = colNameList.ToDictionary(item => new ExcelCellPoint(item.ExcelAddress).Col, item => item);// key是第n列

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList == null
                ? GetExcelListArgs<T>.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;

            //var ctor = type.GetConstructor(new Type[] { });
            //if (ctor == null) throw new ArgumentException($"通过反射无法得到'{type.FullName}'的一个无构造参数的构造器.");

            var dictPropAttrs = new Dictionary<string, Dictionary<string, Attribute>>();//属性里包含的Attribute

            #region 内置的Attribute
            var dictUnique = new Dictionary<string, Dictionary<string, bool>>();//属性的 UniqueAttribute
            var dictKVSet = new Dictionary<string, Dictionary<string, bool>>();//属性的 KVSetAttribute
            string key_UniqueAttribute = typeof(UniqueAttribute).FullName;
            string key_KVSetAttribute = typeof(KVSetAttribute).FullName;

            var cache_PropertyInfo = new Dictionary<string, PropertyInfo>();
            foreach (ExcelCellInfo excelCellInfo in colNameList)
            {
                //int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                //if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                //{
                //    continue;
                //}
                //string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                //if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                var propName = GetPropName<T>(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All, out bool needContinue);
                if (needContinue) continue;

                //if (!cache_PropertyInfo.ContainsKey(propName))
                //{
                //    var pInfo2 = type.GetProperty(propName);
                //    if (pInfo2 == null) //防御式编程判断
                //    {
                //        throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                //    }
                //    cache_PropertyInfo.Add(propName, pInfo2);
                //}

                //PropertyInfo pInfo = cache_PropertyInfo[propName];

                PropertyInfo pInfo = GetPropertyInfo<T>(cache_PropertyInfo, propName, type);

                #region 初始化Attr要处理相关的数据
                dictPropAttrs.Add(pInfo.Name, new Dictionary<string, Attribute>());//这里new 的Dict 的key 代表的是Attribute的FullName

                var uniqueAttrs = ReflectionHelper.GetAttributeForProperty<UniqueAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (uniqueAttrs.Length > 0)
                {
                    dictPropAttrs[pInfo.Name].Add(key_UniqueAttribute, (UniqueAttribute)uniqueAttrs[0]);
                    dictUnique.Add(pInfo.Name, new Dictionary<string, bool>());
                }

                var KVSetAttrs = ReflectionHelper.GetAttributeForProperty<KVSetAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (KVSetAttrs.Length > 0)
                {
                    dictPropAttrs[pInfo.Name].Add(key_KVSetAttribute, (KVSetAttribute)KVSetAttrs[0]);
                    dictKVSet.Add(pInfo.Name, new Dictionary<string, bool>());
                }

                #endregion
            }
            #endregion

            #region 获得 list
            List<T> list = new List<T>();
            int row = rowIndex;
            Exception exception = null;

            int? step = null;
            switch (args.ScanLine)
            {
                case ScanLine.SingleLine:
                    step = 1;
                    break;
                case ScanLine.MergeLine:
                    //while里面动态计算
                    break;
                default:
                    throw new Exception("不支持的ScanLine");
            }

            var excelCellInfoNeedTrim = (args.ReadCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            var excelCellInfoNeedMergeLine = (args.ReadCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            var excelCellInfoNeedToDBC = (args.ReadCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;

            var allException = args.GetList_NeedAllException ? new List<Exception>() : null;

#if DEBUG
            var debugvar_whileCount = 0;
#endif
            Func<object[], object> DeletgateCreateInstance = ExpressionTreeExtensions.BuildDeletgateCreateInstance(type, new Type[0]);

            while (true)//异常或者出现空行,触发break;
            {
#if DEBUG
                debugvar_whileCount++;
#endif
                //判断整行数据是否都没有数据
                bool isNoDataAllColumn = true;

                //Sample02_3,12000的数据
                //T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转  1.2-2.1秒
                //T model = type.CreateInstance<T>();//3秒+
                T model = (T)DeletgateCreateInstance(null); //上面的方法给拆开来 . 1.1-1.4

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    //int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                    //if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                    //{
                    //    continue;
                    //}
                    //string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                    //if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                    var propName = GetPropName<T>(excelCellInfo.ExcelAddress, dictExcelAddressCol, dictExcelColumnIndexToModelPropName_All, out bool needContinue);
                    if (needContinue) continue;

                    //if (!cache_PropertyInfo.ContainsKey(propName))
                    //{
                    //    var pInfo2 = type.GetProperty(propName);
                    //    if (pInfo2 == null)//防御式编程判断
                    //    {
                    //        throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                    //    }
                    //    cache_PropertyInfo.Add(propName, pInfo2);
                    //}

                    //PropertyInfo pInfo = cache_PropertyInfo[propName];

                    PropertyInfo pInfo = GetPropertyInfo<T>(cache_PropertyInfo, propName, type);
                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

#if DEBUG
                    string value;
                    if (pInfo.PropertyType == typeof(DateTime?) || pInfo.PropertyType == typeof(DateTime))
                    {
                        //todo:对于日期类型的,有时候要获取Cell.Value, 有空了修改
                        value = GetMergeCellText(ws, row, col);
                    }
                    else
                    {
                        value = GetMergeCellText(ws, row, col);
                    }
#else
                    string value =  GetMegerCellText(ws, row, col);
#endif

                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);

                    var propAttrs = dictPropAttrs[pInfo.Name];//当前属性的所有特性]

                    if (!valueIsNullOrEmpty)
                    {

                        isNoDataAllColumn = false;

                        #region 判断每个单元格的开头
                        if (args.EveryCellPrefix?.Length > 0)
                        {
                            var indexof = value.IndexOf(args.EveryCellPrefix, StringComparison.Ordinal);
                            if (indexof == -1)
                            {
                                throw new System.ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
                            }
                            value = value.RemovePrefix(args.EveryCellPrefix);
                        }
                        #endregion

                        #region 对每个单元格进行值的替换
                        if (everyCellReplace != null)
                        {
                            foreach (var item in everyCellReplace)
                            {
                                if (!value.Contains(item.Key))
                                {
                                    continue;
                                }
                                var everyCellReplaceOldValue = item.Key;
                                var everyCellReplaceNewValue = item.Value ?? "";
                                if (everyCellReplaceOldValue?.Length > 0)
                                {
                                    value = value.Replace(everyCellReplaceOldValue, everyCellReplaceNewValue);
                                }
                            }
                        }
                        #endregion

                        #region 对每个单元格进行处理

                        if (excelCellInfoNeedTrim)
                        {
                            value = value.Trim();
                        }
                        if (excelCellInfoNeedMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (excelCellInfoNeedToDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion

                        #region 处理内置的Attribute

                        //var propAttrs = dictPropAttrs[pInfo.Name];//当前属性的所有特性

                        if (propAttrs.ContainsKey(key_UniqueAttribute))
                        {
                            #region uniqueAttr的具体实现
                            var uniqueAttr = (UniqueAttribute)propAttrs[key_UniqueAttribute];

                            if (!valueIsNullOrEmpty)
                            {
                                if (!dictUnique[pInfo.Name].ContainsKey(value))
                                {
                                    dictUnique[pInfo.Name].Add(value, default(bool));
                                }
                                else
                                {
                                    string exception_msg = string.IsNullOrEmpty(uniqueAttr.ErrorMessage) ? $@"属性'{pInfo.Name}'的值:'{value}'出现了重复" : uniqueAttr.ErrorMessage;
                                    throw new ArgumentException(exception_msg, pInfo.Name);
                                }
                            }
                            #endregion
                        }
                        if (propAttrs.ContainsKey(key_KVSetAttribute))
                        {
                            var kvsetAttr = (KVSetAttribute)propAttrs[key_KVSetAttribute];

                            var have_kvsource = args.KVSource.ContainsKey(kvsetAttr.Name);
                            if (kvsetAttr.MustInSet && !have_kvsource)
                            {
                                throw new ArgumentException($@"属性'{pInfo.Name}'的值:'{value}'未找到对应的集合列表", pInfo.Name);
                            }

                            object kvsource = args.KVSource[kvsetAttr.Name];
                            var kvsourceType = kvsource.GetType();

                            //var is_kvsourceType = kvsourceType.GetGenericTypeDefinition() == typeof(KVSource<,>);
                            var is_kvsourceType = kvsourceType.HasImplementedRawGeneric(typeof(KvSource<,>));

                            if (is_kvsourceType)
                            {
                                //var kvsourceTypeTKey = kvsourceType.GenericTypeArguments[0];
                                //var kvsourceTypeTValue = kvsourceType.GenericTypeArguments[1];

                                var prop_kvsource = (IKVSource)kvsource;
                                bool inKvSource = prop_kvsource.ContainsKey(value, out object kv_Value);

                                if (!inKvSource && kvsetAttr.MustInSet)
                                {
                                    var msg = string.IsNullOrEmpty(kvsetAttr.ErrorMessage)
                                        ? $@"属性'{pInfo.Name}'值:'{value}'未在'{kvsetAttr.Name}'集合中出现"
                                        : FormatAttributeMsg(pInfo.Name, model, value, kvsetAttr.ErrorMessage, kvsetAttr.Args);// string.Format(kvsetAttr.ErrorMessage, kvsetAttr.Args);
                                    throw new ArgumentException(msg, pInfo.Name);
                                }

                                var typeKVArr = pInfo.PropertyType.GetGenericArguments();
                                var typeKV = typeof(KV<,>).MakeGenericType(typeKVArr);
                                var modelValue = typeKV.GetConstructor(typeKVArr).Invoke(new object[] { value, kv_Value });
                                typeKV.GetProperty("HasValue").SetValue(modelValue, inKvSource);
                                pInfo.SetValue(model, modelValue);
                            }
                        }
                        #endregion
                    }
                    try
                    {
                        //验证特性
                        GetList_ValidAttribute(pInfo, model, value);
                        //赋值, 注:遇到 KV<,> 类型的统一不处理
                        if (!pInfo.PropertyType.HasImplementedRawGeneric(typeof(KV<,>)))
                        {
                            GetList_SetModelValue(pInfo, model, value);
                        }
                    }
                    catch (Exception e)
                    {
                        exception = e is ArgumentException ? new ArgumentException($"无效的单元格:{new ExcelCellAddress(row, col).Address}", e) : e;
                        break;
                    }
                }

                if (isNoDataAllColumn)
                {
                    if (row == rowIndex)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    break; //出现空行,读取模版结束
                }

                //先添加Step
                if (step != null)
                {
                    row += (int)step;
                }
                else
                {
                    string range = ws.MergedCells[row, 1];
                    if (range == null)
                    {
                        row += 1;
                    }
                    else
                    {
                        var ea = new ExcelAddress(range);
                        row += ea.Rows;
                    }
                }
                //在判断异常

                if (exception != null)
                {
                    if (args.GetList_NeedAllException)
                    {
                        allException.Add(exception);
                        exception = null;
                        continue;
                    }
                    else
                    {
                        throw exception;
                    }
                }

                if (args.WhereFilter == null || args.WhereFilter.Invoke(model))
                {
                    list.Add(model);
                }

            }

            var keyWithExceptionMessageStart = "无效的单元格:";
            if (allException != null && allException.Count > 0)
            {
                bool allExceptionIsArgumentException = true;
                var errGroupMsg = new Dictionary<string, List<string>>();

                foreach (var ex in allException)
                {
                    if (!(ex is ArgumentException))
                    {
                        allExceptionIsArgumentException = false;
                        break;
                    }
                    if (!((ArgumentException)ex).Message.StartsWith(keyWithExceptionMessageStart))
                    {
                        allExceptionIsArgumentException = false;
                        break;
                    }

                    var excelCellAddress = ((ArgumentException)ex).Message.RemovePrefix(keyWithExceptionMessageStart);
                    var exceptionMessage = ((ArgumentException)ex).InnerException.Message;
                    if (!errGroupMsg.ContainsKey(exceptionMessage))
                    {
                        errGroupMsg.Add(exceptionMessage, new List<string>());
                    }

                    errGroupMsg[exceptionMessage].Add(excelCellAddress);
                }

                if (!allExceptionIsArgumentException)
                {
                    throw new AggregateException(allException);
                }

                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                foreach (KeyValuePair<string, List<string>> msg in errGroupMsg)
                {
                    sb.Append(msg.Key);
                    sb2.Clear();
                    if (args.GetList_ErrorMessage_OnlyShowColomn)
                    {
                        var cols = new List<string>();
                        foreach (string excelCellAddress in msg.Value)
                        {
                            cols.Add(ExcelCellPoint.R1C1FormulasReverse(new ExcelCellAddress(excelCellAddress).Column));
                        }

                        foreach (var col in cols.Distinct())
                        {
                            sb2.Append(col).Append("列,");
                        }

                    }
                    else
                    {
                        foreach (var excelCellAddress in msg.Value)
                        {
                            sb2.Append(excelCellAddress).Append(",");
                        }
                    }

                    sb2.RemoveLastChar(',');
                    if (sb2.Length > 0)
                    {
                        sb.Append($"({sb2}),");
                    }
                }

                sb.RemoveLastChar(',');
                throw new ArgumentException(sb.ToString());
            }

            #endregion

            return args.HavingFilter == null ? list : list.Where(item => args.HavingFilter.Invoke(item)).ToList();
        }

        private static PropertyInfo GetPropertyInfo<T>(Dictionary<string, PropertyInfo> cache_PropertyInfo, string propName, Type type)
            where T : class, new()
        {
            if (!cache_PropertyInfo.ContainsKey(propName))
            {
                var pInfo2 = type.GetProperty(propName);
                if (pInfo2 == null) //防御式编程判断
                {
                    throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                }

                cache_PropertyInfo.Add(propName, pInfo2);
            }

            PropertyInfo pInfo = cache_PropertyInfo[propName];
            return pInfo;
        }


        private static string GetPropName<T>(ExcelAddress ExcelAddress, Dictionary<ExcelAddress, int> dictExcelAddressCol,
            Dictionary<int, string> dictExcelColumnIndexToModelPropName_All, out bool needContinue) where T : class, new()
        {
            var propName = "";
            int excelCellInfo_ColIndex = dictExcelAddressCol[ExcelAddress];
            if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null) //不存在,跳过
            {
                needContinue = true;
                return propName;
            }
            propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
            needContinue = string.IsNullOrEmpty(propName);
            return propName;
        }

        public static DataTable GetDataTable(GetExcelListArgs<DataRow> args)
        {
            ExcelWorksheet ws = args.ws;
            int rowIndex = args.RowIndex_Data;
            if (rowIndex <= 0)
            {
                throw new ArgumentException($@"数据起始行值'{rowIndex}'错误,值应该大于0");
            }

            int rowIndex_DataName = args.RowIndex_DataName;
            if (rowIndex_DataName <= 0)
            {
                throw new ArgumentException($@"数据起始行的标题行值'{rowIndex_DataName}'错误,值应该大于0");
            }

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, args.DataColStart, args.DataColEnd, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
            if (colNameList.Count == 0)
            {
                throw new Exception("未读取到单元格标题");
            }
            var dictExcelAddressCol = colNameList.ToDictionary(item => item.ExcelAddress, item => new ExcelCellPoint(item.ExcelAddress).Col);

            #region 获得字典

            DataTable dt = new DataTable();

            foreach (var item in colNameList)
            {
                dt.Columns.Add(item.Value.ToString());
            }

            #endregion

            var everyCellReplace = args.UseEveryCellReplace && args.EveryCellReplaceList == null
                ? GetExcelListArgs<DataRow>.EveryCellReplaceListDefault
                : args.EveryCellReplaceList;


            #region 获得 list

            int row = rowIndex;
            int? step = null;
            switch (args.ScanLine)
            {
                case ScanLine.SingleLine:
                    step = 1;
                    break;
                case ScanLine.MergeLine:
                    //while里面动态计算
                    break;
                default:
                    throw new Exception("不支持的ScanLine");
            }

            var excelCellInfoNeedTrim = (args.ReadCellValueOption & ReadCellValueOption.Trim) == ReadCellValueOption.Trim;
            var excelCellInfoNeedMergeLine = (args.ReadCellValueOption & ReadCellValueOption.MergeLine) == ReadCellValueOption.MergeLine;
            var excelCellInfoNeedToDBC = (args.ReadCellValueOption & ReadCellValueOption.ToDBC) == ReadCellValueOption.ToDBC;

            while (true)
            {
                bool isNoDataAllColumn = true;//判断整行数据是否都没有数据
                var dr = dt.NewRow();

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    string propName = excelCellInfo.Value.ToString();

                    if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                    var col = dictExcelAddressCol[excelCellInfo.ExcelAddress];

                    string value = GetMergeCellText(ws, row, col);
                    bool valueIsNullOrEmpty = string.IsNullOrEmpty(value);

                    if (!valueIsNullOrEmpty)
                    {
                        isNoDataAllColumn = false;

                        #region 判断每个单元格的开头
                        if (args.EveryCellPrefix?.Length > 0)
                        {
                            var indexof = value.IndexOf(args.EveryCellPrefix);
                            if (indexof == -1)
                            {
                                throw new System.ArgumentException($"单元格值有误:当前'{new ExcelCellPoint(row, col).R1C1}'单元格的值不是'" + args.EveryCellPrefix + "'开头的");
                            }
                            value = value.RemovePrefix(args.EveryCellPrefix);
                        }
                        #endregion

                        #region 对每个单元格进行值的替换
                        if (everyCellReplace != null)
                        {
                            foreach (var item in everyCellReplace)
                            {
                                if (!value.Contains(item.Key))
                                {
                                    continue;
                                }
                                var everyCellReplaceOldValue = item.Key;
                                var everyCellReplaceNewValue = item.Value ?? "";
                                if (everyCellReplaceOldValue?.Length > 0)
                                {
                                    value = value.Replace(everyCellReplaceOldValue, everyCellReplaceNewValue);
                                }
                            }
                        }
                        #endregion

                        #region 对每个单元格进行处理

                        if (excelCellInfoNeedTrim)
                        {
                            value = value.Trim();
                        }
                        if (excelCellInfoNeedMergeLine)
                        {
                            value = value.MergeLines();
                        }
                        if (excelCellInfoNeedToDBC)
                        {
                            value = value.ToDBC();
                        }

                        #endregion

                    }

                    //赋值
                    dr[propName] = value;
                }

                if (isNoDataAllColumn)
                {
                    if (row == rowIndex)//数据起始行是空行
                    {
                        throw new Exception("不要上传一份空的模版文件");
                    }
                    break; //出现空行,读取模版结束
                }
                //else
                if (args.WhereFilter == null || args.WhereFilter.Invoke(dr))
                {
                    dt.Rows.Add(dr);
                }

                if (step != null)
                {
                    row += (int)step;
                }
                else
                {
                    string range = ws.MergedCells[row, 1];
                    if (range == null)
                    {
                        row += 1;
                    }
                    else
                    {
                        var ea = new ExcelAddress(range);
                        row += ea.Rows;
                    }
                }
            }

            #endregion

            return args.HavingFilter == null ? dt : dt.AsEnumerable().Where(item => args.HavingFilter.Invoke(item)).CopyToDataTable();
        }

        private static string DealMatchingModelException(MatchingModelException matchingModelException)
        {
            //注:这里的仅针对 MatchingModel.eq
            if ((matchingModelException.MatchingModel & MatchingModel.eq) == MatchingModel.eq)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有多提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版提供了多余的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"{excelCellInfo.ExcelAddress}({excelCellInfo.Value}),");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();

            }
            else if ((matchingModelException.MatchingModel & MatchingModel.gt) == MatchingModel.gt)
            {

                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有少提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版多提供了model属性中不存在的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"{excelCellInfo.ExcelAddress}({excelCellInfo.Value}),");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();
            }
            else if ((matchingModelException.MatchingModel & MatchingModel.lt) == MatchingModel.lt)
            {
                if (matchingModelException.ListExcelCellInfoAndModelType == null || matchingModelException.ListExcelCellInfoAndModelType.Count <= 0)
                {
                    return "模版没有多提供列!";
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("模版少提供了model属性中定义的列:");
                foreach (var item in matchingModelException.ListExcelCellInfoAndModelType)
                {
                    foreach (var excelCellInfo in item.ExcelCellInfoList)
                    {
                        sb.Append($@"'{excelCellInfo.Value}',");
                    }
                }
                sb.RemoveLastChar(',');
                sb.Append("!");
                return sb.ToString();

            }
            else if ((matchingModelException.MatchingModel & MatchingModel.neq) == MatchingModel.neq)
            {
                StringBuilder sb = new StringBuilder();
                return sb.ToString();
            }
            else
            {
                throw new Exception($@"参数{nameof(matchingModelException)},不支持的MatchingMode值");
            }

        }

        /// <summary>
        ///  model的哪些属性是在excel中没有定义的 + model中没有定义 
        /// </summary>
        /// <param name="excelColumnIsNotModelProp"></param> 
        /// <param name="type"></param>
        /// <param name="colNameToCellInfo"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static MatchingModelException GetMatchingModelExceptionCase_lt(List<string> excelColumnIsNotModelProp,
            Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
        {
            if (excelColumnIsNotModelProp.Count <= 0)
            {
                return new MatchingModelException { MatchingModel = MatchingModel.lt, ListExcelCellInfoAndModelType = null };
            }

            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
            foreach (var propName in excelColumnIsNotModelProp)
            {
                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType
                {
                    ModelType = type,
                    ExcelCellInfoList = colNameToCellInfo.ContainsKey(propName)
                        ? colNameToCellInfo[propName]
                        : new List<ExcelCellInfo> { new ExcelCellInfo { Value = propName, ExcelAddress = null, WorkSheet = ws } }
                });
            }

            return new MatchingModelException() { MatchingModel = MatchingModel.lt, ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType };
        }

        /// <summary>
        /// excel的哪些列是在Model中定义了却没有(即,model中缺少的列) + model中没有定义 
        /// </summary>
        /// <param name="modelPropNotExistsExcelColumn"></param> 
        /// <param name="type"></param>
        /// <param name="colNameToCellInfo"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static MatchingModelException GetMatchingModelExceptionCase_gt(List<string> modelPropNotExistsExcelColumn, Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
        {

            if (modelPropNotExistsExcelColumn.Count <= 0)
            {
                return new MatchingModelException { MatchingModel = MatchingModel.eq, ListExcelCellInfoAndModelType = null };
            }

            var listExcelCellInfoAndModelType = new List<ExcelCellInfoAndModelType>();
            foreach (var colName in modelPropNotExistsExcelColumn)
            {
                listExcelCellInfoAndModelType.Add(new ExcelCellInfoAndModelType
                {
                    ModelType = type,
                    ExcelCellInfoList = colNameToCellInfo.ContainsKey(colName)
                        ? colNameToCellInfo[colName]
                        : new List<ExcelCellInfo> { new ExcelCellInfo { Value = colName, ExcelAddress = null, WorkSheet = ws } }
                });
            }

            return new MatchingModelException { MatchingModel = MatchingModel.gt, ListExcelCellInfoAndModelType = listExcelCellInfoAndModelType };

        }

        private static MatchingModel GetMatchingModel(
            Dictionary<int, string> dictExcelColumnIndexToExcelColName,
            Dictionary<int, string> dictExcelColumnIndexToModelPropName_All,
            Dictionary<string, bool> dictModelPropNameExistsExcelColumn,
            out List<string> modelPropNotExistsExcelColumn, out List<string> excelColumnIsNotModelProp)
        {
            if (dictExcelColumnIndexToModelPropName_All == null) throw new ArgumentNullException(nameof(dictExcelColumnIndexToModelPropName_All));
            if (dictModelPropNameExistsExcelColumn == null) throw new ArgumentNullException(nameof(dictModelPropNameExistsExcelColumn));

            modelPropNotExistsExcelColumn = new List<string>();//model属性不在excel列中
            excelColumnIsNotModelProp = new List<string>();//excel列不是model属性

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.eq;
            }

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count > 0 && dictModelPropNameExistsExcelColumn.Keys.Count <= 0)
            {
                return MatchingModel.neq | MatchingModel.gt;
            }

            if (dictExcelColumnIndexToModelPropName_All.Keys.Count <= 0 && dictModelPropNameExistsExcelColumn.Keys.Count > 0)
            {
                return MatchingModel.neq | MatchingModel.lt;
            }


            foreach (var excelColumnIndex in dictExcelColumnIndexToModelPropName_All.Keys)
            {
                if (dictExcelColumnIndexToModelPropName_All[excelColumnIndex] == null)
                {
                    modelPropNotExistsExcelColumn.Add(dictExcelColumnIndexToExcelColName[excelColumnIndex]);
                }
            }

            foreach (var modelPropName in dictModelPropNameExistsExcelColumn.Keys)
            {
                if (!dictModelPropNameExistsExcelColumn[modelPropName])
                {
                    excelColumnIsNotModelProp.Add(modelPropName);
                }
            }

            //这里要出重,因为该方法外层的 colNameToCellInfo 对象的类型从 Dictionary<string, ExcelCellInfo>() 改为了 Dictionary<string, List<ExcelCellInfo>>
            if (modelPropNotExistsExcelColumn.Count > 0)
            {
                modelPropNotExistsExcelColumn = modelPropNotExistsExcelColumn.Distinct().ToList();
            }
            if (excelColumnIsNotModelProp.Count > 0)
            {
                excelColumnIsNotModelProp = excelColumnIsNotModelProp.Distinct().ToList();
            }

            if (excelColumnIsNotModelProp.Count == 0 && modelPropNotExistsExcelColumn.Count == 0)
            {
                return MatchingModel.eq;
            }

            if (excelColumnIsNotModelProp.Count > 0 && modelPropNotExistsExcelColumn.Count > 0)
            {
                return MatchingModel.neq;
            }

            if (modelPropNotExistsExcelColumn.Count > 0)//model属性多,即, excel列的数量 比model属性数量少
            {
                return MatchingModel.neq | MatchingModel.gt;
            }

            if (excelColumnIsNotModelProp.Count > 0)
            {
                return MatchingModel.neq | MatchingModel.lt;
            }

            throw new Exception(nameof(GetMatchingModel) + "程序不对,需要修改");
        }

        #endregion

        #region 读取excel配置

        /// <summary>
        /// 从Excel中获取默认的配置信息
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetIndex">第几个workSheet作为模版,从1开始</param>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, int workSheetIndex)
        {
            if (workSheetIndex <= 0) throw new ArgumentOutOfRangeException(nameof(workSheetIndex));
            var worksheet = GetExcelWorksheet(excelPackage, workSheetIndex);
            EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);
            SetConfigBodyFromExcel_OtherPara(config, worksheet);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="config"></param>
        /// <param name="workSheetName"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void SetDefaultConfigFromExcel(ExcelPackage excelPackage, EPPlusConfig config, string workSheetName)
        {
            if (workSheetName == null) throw new ArgumentNullException(nameof(workSheetName));
            var worksheet = GetExcelWorksheet(excelPackage, workSheetName);
            EPPlusHelper.SetDefaultConfigFromExcel(config, worksheet);
            SetConfigBodyFromExcel_OtherPara(config, worksheet);
        }

        /// <summary>
        ///  这个不能在FillData里面算, 会有问题
        /// </summary>
        /// <param name="config"></param>
        /// <param name="worksheet"></param>
        private static void SetConfigBodyFromExcel_OtherPara(EPPlusConfig config, ExcelWorksheet worksheet)
        {
            foreach (var configItem in config.Body.ConfigList)
            {
                var nth = configItem.Nth;

                var allConfig_interval = config.Body[nth].Option.ConfigLine.Count; //配置共计用了多少列, 默认: 1个配置用了1列

                var mergedCellsList = worksheet.MergedCells.ToList();
                foreach (var configCellInfo in config.Body[nth].Option.ConfigLine)
                {
                    if (worksheet.Cells[configCellInfo.Address].Merge) //item.Address  D4
                    {
                        var addressPrecise = EPPlusHelper.GetMergeCellAddressPrecise(worksheet, configCellInfo.Address); //D4:E4格式的
                        allConfig_interval += new ExcelCellRange(addressPrecise).IntervalCol;

                        configCellInfo.FullAddress = addressPrecise;
                        configCellInfo.IsMergeCell = true;

                    }
                    else
                    {
                        var mergeCellAddress = mergedCellsList.Find(a => a.Contains(configCellInfo.Address));
                        if (mergeCellAddress != null)
                        {
                            allConfig_interval += new ExcelCellRange(mergeCellAddress).IntervalCol;
                            configCellInfo.FullAddress = mergeCellAddress;
                            configCellInfo.IsMergeCell = true;
                        }
                        else
                        {
                            configCellInfo.FullAddress = configCellInfo.Address;
                            configCellInfo.IsMergeCell = false;
                        }
                    }
                }

                config.Body[nth].Option.ConfigLineInterval = allConfig_interval;
            }
        }

        /// <summary>
        /// 让 sheet.Cells.Value 强制从A1单元格开始
        /// </summary>
        /// <param name="sheet"></param>
        public static void SetSheetCellsValueFromA1(ExcelWorksheet sheet)
        {
            //让 sheet.Cells.Value 强制从A1单元格开始
            //遇到问题描述:创建一个excel,在C7,C8,C9,10单元格写入一些字符串, sheet.Cells.Value 是object[4,3]的数组, 但我要的是object[10,3]的数组
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value == null)
            {
                cellA1.Value = null;
            }
        }

        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static ExcelCellPoint GetFirstValueCellPoint(ExcelWorksheet ws) => (ExcelCellPoint)GetFirstValueCellInfo<ExcelCellPoint>(ws);
        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static ExcelCellRange GetFirstValueCellRange(ExcelWorksheet ws) => (ExcelCellRange)GetFirstValueCellInfo<ExcelCellRange>(ws);

        /// <summary>
        /// 获得第一个有值的单元格
        /// </summary>
        /// <typeparam name="TOut"></typeparam>
        /// <param name="ws"></param>
        /// <returns></returns>
        private static object GetFirstValueCellInfo<TOut>(ExcelWorksheet ws)
        {
            var returnType = typeof(TOut);
            EPPlusHelper.SetSheetCellsValueFromA1(ws);
            //if (ws.Cells.Value.GetType() == typeof(object[,]))
            //{
            object[,] arr = ws.Cells.Value as object[,];
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;
                    if (arr[i, j].ToString().Length <= 0) continue;
                    if (returnType == typeof(ExcelCellPoint))
                    {
                        var cell = new ExcelCellPoint(i + 1, j + 1);
                        return cell;
                    }
                    else if (returnType == typeof(ExcelCellRange))
                    {
                        var mergeCellAddress = GetMergeCellAddressPrecise(ws, i + 1, j + 1);
                        var cell = new ExcelCellRange(mergeCellAddress);
                        return cell;
                    }
                    else
                    {
                        throw new ArgumentOutOfRangeException(nameof(returnType), $@"不支持的参数{nameof(returnType)}类型:{returnType}");
                    }
                }
            }
            if (returnType == typeof(ExcelCellPoint))
            {
                return new ExcelCellPoint();
            }
            else if (returnType == typeof(ExcelCellRange))
            {
                return new ExcelCellPoint();
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(returnType), $@"不支持的参数{nameof(returnType)}类型:{returnType}");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetDefaultConfigFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            SetSheetCellsValueFromA1(sheet);
            config.Head = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$th") };
            SetConfigBodyFromExcel(config, sheet);
            config.Foot = new EPPlusConfigFixedCells() { ConfigCellList = GetConfigFromExcel(sheet, "$tf") };
        }

        /// <summary>
        /// 设置sheetBody配置
        /// </summary>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetConfigBodyFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");
            var sheetMergedCellsList = sheet.MergedCells.ToList();

            var configLine = new List<List<EPPlusConfigFixedCell>>();
            var configExtra = new List<List<EPPlusConfigFixedCell>>();
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (cellStr.Length < 3) //配置至少有4个字符.所以,4个以下的都可以跳过
                    {
                        continue; //不用""比较,.Length速度比较快
                    }

                    if (!cellStr.StartsWith("$tb")) continue;

                    //  {"L15", "付款对象"}, $tb1
                    string cellPosition = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);

                    string nthStr = RegexHelper.GetFirstNumber(cellStr);
                    int nth = Convert.ToInt32(nthStr);
                    if (cellStr.StartsWith("$tbs")) //模版摘要/汇总等信息单元格
                    {
                        string cellConfigValue = Regex.Replace(cellStr, "^[$]tbs" + nthStr, ""); //$需要转义
                        if (configExtra.Count < nth)
                        {
                            configExtra.Add(new List<EPPlusConfigFixedCell>());
                        }

                        if (configExtra[nth - 1].Find(a => a.ConfigValue == cellConfigValue) !=
                            default(EPPlusConfigFixedCell))
                        {
                            throw new ArgumentException($"Excel文件中的$tbs{nth}部分配置了相同的项:{cellConfigValue}");
                        }

                        configExtra[nth - 1].Add(new EPPlusConfigFixedCell() { Address = cellPosition, ConfigValue = cellConfigValue.Trim(), });
                    }
                    else if (cellStr.StartsWith($"$tb{nthStr}$")) //模版提供了多少行,若没有配置,在调用FillData()时默认提供1行  $tb1$1
                    {
                        string cellConfigValue = Regex.Replace(cellStr, $@"^[$]tb{nth}[$]", ""); //$需要转义, 这个值一般都是数字

                        if (!int.TryParse(cellConfigValue, out int cellConfigValueInt))
                        {
                            if (string.Compare(cellConfigValue, "max", StringComparison.OrdinalIgnoreCase) == 0) //$tb1$max这种配置的
                            {
                                cellConfigValueInt = EPPlusConfig.MaxRow07 - i;
                            }
                            else
                            {
                                throw new Exception("指定提供了多少行的配置项的值无效");
                            }
                        }

                        var nthOption = config.Body[nth].Option;
                        if (nthOption.MapperExcelTemplateLine != null)
                        {
                            throw new ArgumentException($"Excel文件中重复配置了项$tb{nthStr}${cellConfigValue}");
                        }

                        nthOption.MapperExcelTemplateLine = cellConfigValueInt;
                    }
                    else //StartsWith($"$tb{nthStr}")
                    {
                        string cellConfigValue = Regex.Replace(cellStr, "^[$]tb" + nthStr, ""); //$需要转义

                        if (configLine.Count < nth)
                        {
                            configLine.Add(new List<EPPlusConfigFixedCell>());
                        }

                        if (configLine[nth - 1].Find(a => a.ConfigValue == cellConfigValue) != default(EPPlusConfigFixedCell))
                        {
                            throw new ArgumentException($"Excel文件中的$tb{nth}部分配置了相同的项:{cellConfigValue}");
                        }

                        if (sheet.Cells[i + 1, j + 1].Merge)
                        {
                            //只有多行合并时才会影响填充数据时每次新增的行数(多列合并不影响数据的填充)
                            //所以,{"A15:A17", "发生日期"}, 这种情况的key可以写成A15,也可以是A15:K17
                            //{"A15:K17", "发生日期"},只有这种情况的key才要写成A15:K17
                            //后续补充:如果单元格类型是{"A15:A17", "发生日期"},然后key是A15:K17.
                            //在生成excel后打开时会提示***.xlsx中发现不可读取的内容。是否恢复此工作簿的内容.
                            //选择修复文档内容后,里面的内容是正确的(至少我测试的几个是这样的)
                            //所以,同行多列合并的单元格的key 必须是 A15 这种格式的
                            var newKey = sheetMergedCellsList.Find(a => a.Contains(cellPosition));
                            if (newKey == null)
                            {
                                //描述出现null的情况
                                /*
                                 * F10 G10
                                 * F11 G11
                                 * F12 G12
                                 * 这些单元格被合并为一个单元格,即用F10:G12来描述
                                 * 此时,配置单元格读取应该是F10,G10将不会被读取,
                                 * 直到上面为止,都是正确的,但是,偏偏有一个神一样的操作,
                                 * 当excel模版出现不规范操作(Excel一眼看上去将没有问题),G10单元格被读取出来后,那么在sheetMergedCellsList中肯定找不到
                                 * 然后下面一行代码就抛出未将对象引用设置到对象的实例异常
                                 * 该操作是:B10, D10, F10, G10单元格均有配置项,B10:C12进行单元格合并,然后用格式刷,对D10:E12, F10:G10进行格式化
                                 * 结果就是G10的配置项将会被隐藏,excel会提示 合并单元格时，仅保留左上角的值，而放弃其他值。但是这个其他值没有被清空,而是看不到了
                                 * 如果手动的合并F10:G10,Excel将会alert此操作会仅保留左上角的值
                                 */
                                throw new Exception($"工作簿{sheet.Name}的单元格{cellPosition}存在配置问题,请检查({cellPosition}是合并单元格,请取消合并,并把单元格的值给清空,然后重新合并)");
                            }

                            var cells = newKey.Split(':');

                            // ReSharper disable once ConvertIfStatementToConditionalTernaryExpression
                            if (RegexHelper.GetFirstNumber(cells[0]) == RegexHelper.GetFirstNumber(cells[1])) //是同一行的
                            {
                                configLine[nth - 1].Add(new EPPlusConfigFixedCell() { Address = cellPosition, ConfigValue = cellConfigValue });
                            }
                            else
                            {
                                configLine[nth - 1].Add(new EPPlusConfigFixedCell() { Address = newKey, ConfigValue = cellConfigValue });
                            }
                        }
                        else
                        {
                            configLine[nth - 1].Add(new EPPlusConfigFixedCell() { Address = cellPosition, ConfigValue = cellConfigValue });
                        }
                    }

                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉. 注用这种方式清空值,,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                    sheet.Cells[i + 1, j + 1].Value = null; //修复bug:当只有一个配置时,这个deleteLastSpaceLine 为false,然后在excel筛选的时候能出来一行空白(后期已经修复)
                }
            }

            for (int i = 0; i < configLine.Count; i++)
            {
                config.Body[i + 1].Option.ConfigLine = configLine[i];
            }

            for (int i = 0; i < configExtra.Count; i++)
            {
                config.Body[i + 1].Option.ConfigExtra = configExtra[i];
            }

        }

        /// <summary>
        /// 设置sheetFoot配置
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startWith"></param>
        /// <returns></returns>
        private static List<EPPlusConfigFixedCell> GetConfigFromExcel(ExcelWorksheet sheet, string startWith)
        {
            if (!startWith.StartsWith("$")) throw new ArgumentException("配置项必须是$开头");

            object[,] arr = sheet.Cells.Value as object[,];
            Debug.Assert(arr != null, nameof(arr) + " != null");

            var fixedCellsInfoList = new List<EPPlusConfigFixedCell>();
            var replaceStr = startWith.RemovePrefix("$");
            for (var i = 0; i < arr.GetLength(0); i++)
            {
                for (var j = 0; j < arr.GetLength(1); j++)
                {
                    if (arr[i, j] == null) continue;

                    string cellStr = arr[i, j].ToString().Trim();
                    if (!cellStr.StartsWith(startWith)) continue;

                    // {"G6", "公司名称"},
                    string key = ExcelCellPoint.R1C1FormulasReverse(j + 1) + (i + 1);
                    string val = Regex.Replace(cellStr, $"^[$]{replaceStr}", "").Trim(); //$需要转义
                    if (fixedCellsInfoList.Find(a => a.ConfigValue == val) != null)
                    {
                        throw new ArgumentException($"Excel文件中的{startWith}部分配置了相同的项:{val}");
                    }

                    fixedCellsInfoList.Add(new EPPlusConfigFixedCell() { Address = key, ConfigValue = val.Trim() });
                    //arr[i,j] = "";//把当前单元格值清空
                    //sheet.Cells[i + 1, j + 1].Value = ""; //不知道为什么上面的清空不了,但是有时候有能清除掉 注用这种方式清空值,,每个单元格 会有一个 ascii 为 9 (\t) 的符号进去
                    sheet.Cells[i + 1, j + 1].Value = null; //统一用 null 来清空单元格
                }
            }
            return fixedCellsInfoList;
        }

        #endregion

    }
}
