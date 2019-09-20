using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Diagnostics;
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

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, 1, EPPlusConfig.MaxCol07, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
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
                        PropertyInfo pInfoTemp = null;
                        var propNameTemp = dictExcelColumnIndexToModelPropName_Temp[excelColumnIndex];
                        //不做属性的 DisplayExcelColumnName = 当前属性的验证 (因为还没想到这个属性是一定要验证的情况)
                        //if (dictModelPropNameToExcelColumnName.ContainsKey(propNameTemp) && dictModelPropNameToExcelColumnName[propNameTemp] == propName)
                        //{
                        //   pInfoTemp = type.GetProperty(propName);
                        //}

                        pInfoTemp = type.GetProperty(propNameTemp);
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

            var ctor = type.GetConstructor(new Type[] { });
            if (ctor == null) throw new ArgumentException($"通过反射无法得到'{type.FullName}'的一个无构造参数的构造器.");

            var dictPropAttrs = new Dictionary<string, Dictionary<string, Attribute>>();//属性里包含的Attribute

            //内置的Attribute
            var dictUnique = new Dictionary<string, Dictionary<string, bool>>();//属性的 UniqueAttribute
            string key_UniqueAttribute = typeof(UniqueAttribute).FullName;

            foreach (ExcelCellInfo excelCellInfo in colNameList)
            {
                int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                {
                    continue;
                }
                string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                PropertyInfo pInfo = type.GetProperty(propName);
                if (pInfo == null)//防御式编程判断
                {
                    throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                }

                #region 初始化Attr要处理相关的数据
                var attrDict = new Dictionary<string, Attribute>() { };//key 是Attribute的FullName
                dictPropAttrs.Add(pInfo.Name, attrDict);

                var uniqueAttrs = ReflectionHelper.GetAttributeForProperty<UniqueAttribute>(pInfo.DeclaringType, pInfo.Name);
                if (uniqueAttrs.Length > 0)
                {
                    dictPropAttrs[pInfo.Name].Add(key_UniqueAttribute, (UniqueAttribute)uniqueAttrs[0]);
                    dictUnique.Add(pInfo.Name, new Dictionary<string, bool>());
                }

                #endregion
            }

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
            while (true)
            {
                bool isNoDataAllColumn = true;//判断整行数据是否都没有数据
                T model = ctor.Invoke(new object[] { }) as T; //返回的是object,需要强转

                foreach (ExcelCellInfo excelCellInfo in colNameList)
                {
                    int excelCellInfo_ColIndex = dictExcelAddressCol[excelCellInfo.ExcelAddress];
                    if (dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex] == null)//不存在,跳过
                    {
                        continue;
                    }
                    string propName = dictExcelColumnIndexToModelPropName_All[excelCellInfo_ColIndex];
                    if (string.IsNullOrEmpty(propName)) continue;//理论上,这种情况不存在,即使存在了,也要跳过

                    PropertyInfo pInfo = type.GetProperty(propName);
                    if (pInfo == null)//防御式编程判断
                    {
                        throw new ArgumentException($@"Type:'{type}'的property'{propName}'未找到");
                    }

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

                        if (dictPropAttrs[pInfo.Name].ContainsKey(key_UniqueAttribute))
                        {
                            var uniqueAttr = (UniqueAttribute)dictPropAttrs[pInfo.Name][key_UniqueAttribute];
                            var uniqueAttrAttr_ErrorMsg_IsNullOrEmpty = string.IsNullOrEmpty(uniqueAttr.ErrorMessage);
                            if (!valueIsNullOrEmpty)
                            {
                                if (!dictUnique[pInfo.Name].ContainsKey(value))
                                {
                                    dictUnique[pInfo.Name].Add(value, default(bool));
                                }
                                else
                                {
                                    string exception_msg = uniqueAttrAttr_ErrorMsg_IsNullOrEmpty ? $@"属性'{pInfo.Name}'的值:'{value}'出现了重复" : uniqueAttr.ErrorMessage;
                                    throw new ArgumentException(exception_msg, pInfo.Name);
                                }
                            }

                        }
                        #endregion
                    }
                    try
                    {
                        //验证特性
                        GetList_ValidAttribute(pInfo, model, value);
                        //赋值
                        GetList_SetModelValue(pInfo, model, value);
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

            var colNameList = GetExcelColumnOfModel(ws, rowIndex_DataName, 1, EPPlusConfig.MaxCol07, args.POCO_Property_AutoRename_WhenRepeat, args.POCO_Property_AutoRenameFirtName_WhenRepeat);
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
        private static MatchingModelException GetMatchingModelExceptionCase_gt(List<string> modelPropNotExistsExcelColumn,
            Type type, Dictionary<string, List<ExcelCellInfo>> colNameToCellInfo, ExcelWorksheet ws)
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
                        var addressPrecise =EPPlusHelper.GetMergeCellAddressPrecise(worksheet, configCellInfo.Address); //D4:E4格式的
                        allConfig_interval += new ExcelCellRange(addressPrecise).IntervalCol;

                        configCellInfo.FullAddress= addressPrecise;
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
        /// 
        /// </summary>
        /// <param name="config"></param>
        /// <param name="sheet"></param>
        public static void SetDefaultConfigFromExcel(EPPlusConfig config, ExcelWorksheet sheet)
        {
            //让 sheet.Cells.Value 强制从A1单元格开始
            //遇到问题描述:创建一个excel,在C7,C8,C9,10单元格写入一些字符串, sheet.Cells.Value 是object[4,3]的数组, 但我要的是object[10,3]的数组
            var cellA1 = sheet.Cells[1, 1];
            if (!cellA1.Merge && cellA1.Value == null)
            {
                cellA1.Value = null;
            }

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

        #region 一些帮助方法

        /// <summary>
        /// 获得精确的合并单元格地址
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, int row, int col)
        {
            var mergeCellAddress = ws.MergedCells[row, col];//最准确的合并单元格值
            if (mergeCellAddress == null)
            {
                //不是合并单元格
                return new ExcelCellPoint(row, col).R1C1;
            }
            else
            {
                return mergeCellAddress;
            }
        }

        public static string GetMergeCellAddressPrecise(ExcelWorksheet ws, string r1c1)
        {
            var excelRange = new ExcelCellRange(r1c1);
            if (excelRange.End.Col == 0) //r1c1 为 D4  这种值
            {
                return GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.Start.Col);
            }
            else
            {
                return GetMergeCellAddressPrecise(ws, excelRange.Start.Row, excelRange.End.Col);
            }
        }



        public static string GetLeftCellAddress(ExcelWorksheet ws, string currentCellAddress)
        {
            var ea = new ExcelAddress(currentCellAddress);
            var row = ea.Start.Row;
            var col = ea.Start.Column;
            var cell = ws.Cells[row, col];
            if (cell.Merge)
            {
                var mergeCellAddress = ws.MergedCells[row, col];//最准确的合并单元格值
                var mergeCell = new ExcelAddress(mergeCellAddress);
                //var margeCell_Range = new ExcelCellRange(mergeCellAddress);

                var leftCellRow = mergeCell.Start.Row;
                var leftCellCol = mergeCell.Start.Column - 1;
                var leftCellAddress = ws.MergedCells[leftCellRow, leftCellCol];
                if (leftCellAddress == null) //左边的单元格是普通的单元格
                {
                    return new ExcelCellPoint(leftCellRow, leftCellCol).R1C1;
                }
                else
                {
                    ea = new ExcelAddress(leftCellAddress);
                    return ea.Address;
                }
            }
            else
            {
                return new ExcelCellPoint(row, col - 1).R1C1;
            }
        }

        /// <summary>
        /// 获得合并单元格的值 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string GetMergeCellText(ExcelWorksheet ws, int row, int col)
        {
            string mergeCellAddress = ws.MergedCells[row, col];
            if (mergeCellAddress == null) return GetCellText(ws, row, col); //不是合并单元格
            var ea = new ExcelAddress(mergeCellAddress);
            return ws.Cells[ea.Start.Row, ea.Start.Column].Text;
        }

        /// <summary>
        /// 如果是合并单元格,请传入第一个合并单元格的坐标
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">从1开始</param>
        /// <param name="col">从1开始</param>
        /// <param name="when_TextProperty_NullReferenceException_Use_ValueProperty"></param>
        /// <returns></returns>
        public static string GetCellText(ExcelWorksheet ws, int row, int col, bool when_TextProperty_NullReferenceException_Use_ValueProperty = true)
        {
            return GetCellText(ws.Cells[row, col], when_TextProperty_NullReferenceException_Use_ValueProperty);
        }

        public static string GetCellText(ExcelRange cell, bool when_TextProperty_NullReferenceException_Use_ValueProperty = true)
        {
            //if (cell.Merge) throw new Exception("没遇到过这个情况的");
            //return cell.Text; //这个没有科学计数法  注:Text是Excel显示的值,Value是实际值.
            try
            {
                return cell.Text;//有的单元格通过cell.Text取值会发生异常,但cell.Value却是有值的
            }
            catch (System.NullReferenceException)
            {
                if (when_TextProperty_NullReferenceException_Use_ValueProperty)
                {
                    return Convert.ToString(cell.Value);
                }
                throw;
            }
        }

        /// <summary>
        /// 根据dict检查指定单元格的值是否符合预先定义.
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="dict">k:r1c1, v:具体值</param>

        public static bool CheckWorkSheetCellValue(ExcelWorksheet ws, Dictionary<string, string> dict)
        {
            //var dict = new Dictionary<string, string>() { { "A1", "序号" } };
            foreach (var key in dict.Keys)
            {
                var cell = new ExcelCellPoint(key);
                if (EPPlusHelper.GetCellText(ws, cell.Row, cell.Col) != dict[key])
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 冻结窗口面板
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row">冻结[1,Row)行</param>
        /// <param name="column">冻结{1,Column) 列</param>
        public static void FreezePanes(ExcelWorksheet ws, int row, int column)
        {
            ws.View.FreezePanes(row, column);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="sheetTitleLineNumber">工作簿标题行,key:工作簿名字,value:行号</param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, Dictionary<string, int> sheetTitleLineNumber)
        {
            if (sheetTitleLineNumber == null)
            {
                sheetTitleLineNumber = new Dictionary<string, int>();
            }
            ExcelWorksheets wss = excelPackage.Workbook.Worksheets;
            List<DefaultConfig> list = new List<DefaultConfig>();
            foreach (var ws in wss)
            {
                int titleLine = sheetTitleLineNumber != null || sheetTitleLineNumber.ContainsKey(ws.Name)
                    ? sheetTitleLineNumber[ws.Name]
                    : 2;
                list.Add(FillExcelDefaultConfig(ws, titleLine));
            }

            return list;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileOutDirectoryName"></param>
        /// <param name="sheetTitleLineNumber"></param>
        /// <param name="cellCustom"></param>
        /// <returns></returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(string filePath, string fileOutDirectoryName, Dictionary<int, int> sheetTitleLineNumber = null, Action<ExcelRange> cellCustom = null)
        {
            List<DefaultConfig> defaultConfigList;
            using (MemoryStream ms = new MemoryStream())
            //using (FileStream fs = System.IO.File.OpenRead(filePath))
            using (FileStream fs = new System.IO.FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ExcelPackage excelPackage = new ExcelPackage(fs))
            {
                defaultConfigList = FillExcelDefaultConfig(excelPackage, sheetTitleLineNumber, cellCustom);
                excelPackage.SaveAs(ms);
                ms.Position = 0;
                ms.Save($@"{fileOutDirectoryName}\{Path.GetFileNameWithoutExtension(filePath)}_Result.xlsx");
            }
            return defaultConfigList;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="sheetTitleLineNumber">工作簿标题行,key:第几个工作簿,从1开始,value:行号</param>
        /// <param name="cellCustom"></param>
        /// <returns>工作簿Name,DatTable的创建代码</returns>
        public static List<DefaultConfig> FillExcelDefaultConfig(ExcelPackage excelPackage, Dictionary<int, int> sheetTitleLineNumber, Action<ExcelRange> cellCustom = null)
        {
            ExcelWorksheets wss = excelPackage.Workbook.Worksheets;
            List<DefaultConfig> list = new List<DefaultConfig>();
            var eachCount = 0;
            foreach (var ws in wss)
            {
                int titleLine = sheetTitleLineNumber == null
                    ? 1
                    : sheetTitleLineNumber.ContainsKey(eachCount) ? sheetTitleLineNumber[eachCount] : 1;
                list.Add(FillExcelDefaultConfig(ws, titleLine, cellCustom));
                eachCount++;
            }
            return list;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="titleLineNumber"></param>
        /// <param name="cellCustom">对单元格进行额外处理</param>
        /// <returns></returns>
        public static DefaultConfig FillExcelDefaultConfig(ExcelWorksheet ws, int titleLineNumber, Action<ExcelRange> cellCustom = null)
        {
            var colNameList = new List<ExcelCellInfoValue>();
            var nameRepeatCounter = new Dictionary<string, int>();
            #region 获得colNameList

            int col = 1;
            while (col <= EPPlusConfig.MaxCol07)
            {
                var excelColName = ws.Cells[titleLineNumber, col].Merge ? GetMergeCellText(ws, titleLineNumber, col) : GetCellText(ws, titleLineNumber, col);

                var destColVal = ExtractName(excelColName).Trim().MergeLines();
                if (string.IsNullOrEmpty(destColVal))
                {
                    break;
                }

                var thisColNameInfo = new ExcelCellInfoValue()
                {
                    Name = destColVal,
                    ExcelColNameIndex = col,
                    //ExcelColName = excelColName,
                    ExcelColName = excelColName.Trim().MergeLines(),
                };

                AutoRename(colNameList, nameRepeatCounter, thisColNameInfo, true);

                if (ws.Cells[titleLineNumber, col].Merge)
                {
                    var address = ws.MergedCells[titleLineNumber, col];
                    var range = new ExcelCellRange(address);
                    col += range.IntervalCol + 1;
                }
                else
                {
                    col++;
                }
            }

            #endregion

            #region 给单元格赋值

            int fillBodyLine;
            if (ws.Cells[titleLineNumber, 1].Merge)
            {
                var address = ws.MergedCells[titleLineNumber, 1];
                var range = new ExcelCellRange(address);
                fillBodyLine = range.Start.Row + range.IntervalRow + 1;
            }
            else
            {
                fillBodyLine = titleLineNumber + 1;
            }

            for (int i = 0; i < colNameList.Count; i++)
            {
                ws.Cells[fillBodyLine, colNameList[i].ExcelColNameIndex].Value = $@"$tb1{(colNameList[i].IsRename ? colNameList[i].NameNew : colNameList[i].Name)}";
                cellCustom?.Invoke(ws.Cells[titleLineNumber + 1, i + 1]);
            }

            #endregion

            #region sb_CreateClassSnippet + sb_CreateDateTableSnippet
            StringBuilder sb_CreateClassSnippet = new StringBuilder();
            sb_CreateClassSnippet.AppendLine($"public class {ws.Name} {{");

            StringBuilder sb_CreateDateTableSnippet = new StringBuilder();
            sb_CreateDateTableSnippet.AppendLine($@"DataTable dt = new DataTable();");
            StringBuilder sbColumn = new StringBuilder();
            StringBuilder sbAddDr = new StringBuilder();
            StringBuilder sbColumnType = new StringBuilder();
            sbAddDr.AppendLine($@"//var dr = dt.NewRow();");

            #region 关键字
            var columnTypeList_DateTime = new List<string>()
            {
                "时间", "日期", "date", "time","今天","昨天","明天","前天","day"
            };
            var columnTypeList_String = new List<string>()
            {
                "id","身份证","银行卡","卡号","手机","mobile","tel","序号","number","编号","No"
            };
            #endregion

            #region 关键字tolower
            for (int i = 0; i < columnTypeList_DateTime.Count; i++)
            {
                columnTypeList_DateTime[i] = columnTypeList_DateTime[i].ToLower();
            }
            for (int i = 0; i < columnTypeList_String.Count; i++)
            {
                columnTypeList_String[i] = columnTypeList_String[i].ToLower();
            }
            #endregion

            foreach (var colName in colNameList)
            {
                var propName = colName.IsRename ? colName.NameNew : colName.Name;
                sbColumn.AppendLine($"dt.Columns.Add(\"{propName}\");");
                sbAddDr.AppendLine($"//dr[\"{propName}\"] = ");

                var propName_lower = propName.ToLower();
                bool sb_CrateClassSnippe_AppendLine_InForeach = false;


                if (colName.IsRename)
                {
                    sb_CreateClassSnippet.AppendLine($" [ExcelColumnIndex({colName.ExcelColNameIndex})]");
                    sb_CreateClassSnippet.AppendLine($" [DisplayExcelColumnName(\"{colName.ExcelColName}\")]");
                }

                if (colName.ExcelColName != colName.Name)
                {
                    if (!colName.IsRename)//上面添加过了,这里不在添加
                    {
                        sb_CreateClassSnippet.AppendLine($" [DisplayExcelColumnName(\"{colName.ExcelColName}\")]");
                    }
                }
                foreach (var item in columnTypeList_DateTime)
                {
                    if (propName_lower.IndexOf(item, StringComparison.Ordinal) != -1)
                    {
                        sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(DateTime);");
                        sb_CreateClassSnippet.AppendLine($" public DateTime {propName} {{ get; set; }}");
                        sb_CrateClassSnippe_AppendLine_InForeach = true;
                        break;
                    }
                }

                foreach (var item in columnTypeList_String)
                {
                    if (propName_lower.IndexOf(item, StringComparison.Ordinal) != -1)
                    {
                        sbColumnType.AppendLine($"dt.Columns[\"{propName}\"].DataType = typeof(String);");
                        sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                        sb_CrateClassSnippe_AppendLine_InForeach = true;
                        break;//处理过了就break,不然会重复处理 譬如 银行卡号, 此时符合 银行卡 和卡号
                    }
                }

                if (!sb_CrateClassSnippe_AppendLine_InForeach)
                {
                    sb_CreateClassSnippet.AppendLine($" public string {propName} {{ get; set; }}");
                }

            }
            sb_CreateDateTableSnippet.Append(sbColumn);
            sb_CreateDateTableSnippet.Append(sbColumnType);
            sbAddDr.AppendLine("//dt.Rows.Add(dr);");
            sb_CreateDateTableSnippet.Append(sbAddDr);

            sb_CreateClassSnippet.AppendLine("}");
            #endregion

            return new DefaultConfig()
            {
                WorkSheetName = ws.Name,
                CrateDataTableSnippe = sb_CreateDateTableSnippet.ToString(),
                CrateClassSnippe = sb_CreateClassSnippet.ToString(),
                ClassPropertyList = colNameList
            };

        }

        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        private static void AutoRename(List<string> nameList, Dictionary<string, int> nameRepeatCounter, string name, bool renameFirstNameWhenRepeat)
        {

            if (!nameRepeatCounter.ContainsKey(name))
            {
                nameRepeatCounter.Add(name, 0);
            }

            if (!nameList.Contains(name) && nameRepeatCounter[name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name] == 1)
                {
                    for (int i = 0; i < nameList.Count; i++)
                    {
                        if (nameList[i] == name)
                        {
                            nameList[i] = nameList[i] + "1";
                            break;
                        }
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name];
                nameList.Add($@"{name}{++currentCounterVal}");
            }

            nameRepeatCounter[name] = ++nameRepeatCounter[name];
        }


        /// <summary>
        /// 自动重命名
        /// </summary>
        /// <param name="nameList">重名后的name集合</param>
        /// <param name="nameRepeatCounter">name重复的次数</param>
        /// <param name="name">要传入的name值</param>
        /// <param name="renameFirstNameWhenRepeat">当重名时,重命名第一个名字</param>
        private static void AutoRename(List<ExcelCellInfoValue> nameList, Dictionary<string, int> nameRepeatCounter, ExcelCellInfoValue name, bool renameFirstNameWhenRepeat)
        {
            if (!nameRepeatCounter.ContainsKey(name.Name))
            {
                nameRepeatCounter.Add(name.Name, 0);
            }

            if (nameList.Find(a => a.Name == name.Name) == null && nameRepeatCounter[name.Name] == 0)
            {
                nameList.Add(name);
            }
            else
            {
                //如果出现重复,把第一个名字添加后缀1
                if (renameFirstNameWhenRepeat && nameRepeatCounter[name.Name] == 1)
                {
                    foreach (var t in nameList)
                    {
                        if (t.Name != name.Name) continue;
                        t.IsRename = true;
                        t.NameNew = t.Name + "1";
                        break;
                    }
                }
                //必须要先用一个变量保存,使用 ++colNames_Counter[destColVal] 会把 colNames_Counter[destColVal] 值变掉
                var currentCounterVal = nameRepeatCounter[name.Name];
                name.IsRename = true;
                name.NameNew = $@"{name.Name}{++currentCounterVal}";
                nameList.Add(name);
            }
            nameRepeatCounter[name.Name] = ++nameRepeatCounter[name.Name];
        }

        /// <summary>
        /// 获得excel填写的配置内容
        /// </summary>
        /// <param name="content"></param>
        /// <param name="outResultPrefix"></param>
        /// <returns></returns>
        public static string GetFillDefaultConfig(string content, string outResultPrefix = "$tb1")
        {
            if (string.IsNullOrEmpty(content)) return content;
            content = content.TrimEnd();
            content.RemoveLastChar('\n');//excel选择列复制出来到文本上有换行,最后一个字符的ascii 是10 \n
            content.RemoveLastChar('\r');//如果是自己敲入的回车,那么也去掉
            var excel_cell_split = new char[] { '	', ' ', };// 两个单元格之间间隔的符号(\t),空格
            string[] splits = content.Split(excel_cell_split, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();
            StringBuilder sbColumn = new StringBuilder();
            foreach (var item in splits)
            {
                var newName = ExtractName(item);
                sb.Append($@"{outResultPrefix}{newName}{excel_cell_split[0]}");
                sbColumn.AppendLine($"dt.Columns.Add(\"{newName}\");");
            }

            sb.RemoveLastChar(excel_cell_split[0]);

            //sb.AppendLine().AppendLine().AppendLine();
            //sb.AppendLine($@"DataTable dt = new DataTable();");
            //sb.Append(sbColumn.ToString());

            return sb.ToString();
        }


        #region 获得单元格
        /// <summary>
        /// 根据值获的excel中对应的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static List<ExcelCellInfo> GetCellsBy(ExcelWorksheet ws, string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            object[,] cellsValue = ws.Cells.Value as object[,];
            if (cellsValue == null) throw new ArgumentNullException();
            return GetCellsBy(ws, cellsValue, a => a != null && a.ToString() == value);
        }

        /// <summary>
        /// 根据值获的excel中对应的单元格
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="cellsValue">一般通过ws.Cells.Value as object[,] 获得 </param>
        /// <param name="match">示例: a => a != null &amp;&amp; a.GetType() == typeof(string) &amp;&amp; ((string) a == "备注")</param>
        /// <returns></returns>
        public static List<ExcelCellInfo> GetCellsBy(ExcelWorksheet ws, object[,] cellsValue, Predicate<object> match)
        {
            if (cellsValue == null) throw new ArgumentNullException(nameof(cellsValue));

            var result = new List<ExcelCellInfo>();
            for (int i = 0; i < cellsValue.GetLength(0); i++)
            {
                for (int j = 0; j < cellsValue.GetLength(1); j++)
                {
                    if (match != null && match.Invoke(cellsValue[i, j]))
                    {
                        result.Add(new ExcelCellInfo
                        {
                            WorkSheet = ws,
                            ExcelAddress = new ExcelAddress(i + 1, j + 1, i + 1, j + 1),
                            Value = cellsValue[i, j]
                        });
                    }
                }
            }

            return result;
        }
        #endregion

        #region 科学计数法的cell转成文本格式的cell

        /// <summary>
        /// 科学计数法的cell转成文本格式的cell
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="savePath"></param>
        /// <param name="containNoMatchCell">包含不匹配的单元格(即把所有的单元格变成文本格式),true:是.false:仅针对显示成科学计数法的cell变成文本</param>
        /// <returns>是否有进行科学技术法的cell转换.true:是,false:否</returns>
        public static bool ScientificNotationFormatToString(ExcelPackage excelPackage, string savePath, bool containNoMatchCell = false)
        {
            long modifyCellCount = 0;//统计修改的次数

            //下面代码之所以用if-else,是因为这样可以减少很多判断次数.
            if (containNoMatchCell)
            {
                for (int workSheetIndex = 1; workSheetIndex <= excelPackage.Workbook.Worksheets.Count; workSheetIndex++)
                {
                    var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
                    object[,] arr = sheet.Cells.Value as object[,];

                    Debug.Assert(arr != null, nameof(arr) + " != null");
                    for (int i = 0; i < arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < arr.GetLength(1); j++)
                        {
                            if (arr[i, j] != null)
                            {
                                modifyCellCount++;
                                sheet.Cells[i + 1, j + 1].Value = sheet.Cells[i + 1, j + 1].Text;
                            }
                        }
                    }
                }
            }
            else
            {
                for (int workSheetIndex = 1; workSheetIndex <= excelPackage.Workbook.Worksheets.Count; workSheetIndex++)
                {
                    var sheet = GetExcelWorksheet(excelPackage, workSheetIndex);
                    object[,] arr = sheet.Cells.Value as object[,];

                    Debug.Assert(arr != null, nameof(arr) + " != null");
                    for (int i = 0; i < arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < arr.GetLength(1); j++)
                        {
                            if (arr[i, j] != null)
                            {
                                //两段代码唯一的区别
                                var cell = sheet.Cells[i + 1, j + 1];
                                if (cell.Value is double && cell.Text.Length > 11)
                                {
                                    modifyCellCount++;
                                    cell.Value = cell.Text;
                                }
                            }
                        }
                    }
                }
            }

            using (var ms = new MemoryStream())
            {
                excelPackage.SaveAs(ms); // 导入数据到流中 
                ms.Position = 0;
                File.Delete(savePath); //删除文件。如果文件不存在,也不报错
                ms.Save(savePath);
            }

            return modifyCellCount > 0;
        }

        /// <summary>
        /// 处理Excel,将包含科学计数法的cell转成文本格式的cell
        /// </summary>
        /// <param name="fileFullPath">文件路径</param>
        /// <param name="fileSaveAsPath">文件另存为路径</param>
        /// <param name="containNoMatchCell">包含不匹配的单元格(即把所有的单元格变成文本格式),true:是.false:仅针对显示成科学计数法的cell变成文本</param>
        /// <returns>是否有进行科学技术法的cell转换.true:是,false:否</returns>
        public static bool ScientificNotationFormatToString(string fileFullPath, string fileSaveAsPath, bool containNoMatchCell = false)
        {
            using (var fs = File.OpenRead(fileFullPath))
            using (var excelPackage = new ExcelPackage(fs))
            {
                return ScientificNotationFormatToString(excelPackage, fileSaveAsPath, containNoMatchCell);
            }
        }

        #endregion

        #endregion

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
