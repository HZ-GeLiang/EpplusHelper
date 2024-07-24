using EPPlusExtensions.Utils;
using System.Data;

namespace EPPlusExtensions
{
    /// <summary>
    /// 配置信息的数据源
    /// </summary>
    public class EPPlusConfigSource
    {
        public EPPlusConfigSourceHead Head { get; set; } //= new EPPlusConfigSourceHead();
        public EPPlusConfigSourceBody Body { get; set; } //= new EPPlusConfigSourceBody();
        public EPPlusConfigSourceFoot Foot { get; set; } //= new EPPlusConfigSourceFoot();
    }

    public class EPPlusConfigSourceFixedCell<TValue>
    {
        /// <summary>
        /// 单元格配置的值:如 Name
        /// </summary>
        public string ConfigValue { get; set; }

        /// <summary>
        /// 填写的值:如 张三
        /// </summary>
        public TValue FillValue { get; set; }
    }

    public class EPPlusConfigSourceFixedCell
    {
        /// <summary>
        /// 单元格配置的值:如 Name
        /// </summary>
        public string ConfigValue { get; set; }

        /// <summary>
        /// 填写的值:如 张三
        /// </summary>
        public object FillValue { get; set; }
    }

    public class EPPlusConfigSourceConfigExtras<TValue>
    {
        public List<EPPlusConfigSourceFixedCell<TValue>> CellsInfoList { get; set; } = null;

        public EPPlusConfigSourceFixedCell<TValue> this[string key]
        {
            get
            {
                var cell = GetCellAndTryAdd(key);
                return cell;
            }
            set
            {
                var cell = GetCellAndTryAdd(key);
                cell.FillValue = ValueConvertUtil.ConvertToTValue<TValue>(value);
            }
        }

        private EPPlusConfigSourceFixedCell<TValue> GetCellAndTryAdd(string key)
        {
            if (string.IsNullOrEmpty(key)) throw new ArgumentNullException($"{nameof(key)}不能为空");
            if (CellsInfoList is null)
            {
                CellsInfoList = new List<EPPlusConfigSourceFixedCell<TValue>>();
            }

            var cell = CellsInfoList.Find(a => a.ConfigValue == key);
            if (cell is null)
            {
                cell = new EPPlusConfigSourceFixedCell<TValue>() { ConfigValue = key };
                CellsInfoList.Add(cell);
            }

            return cell;
        }

        public static List<EPPlusConfigSourceFixedCell<TValue>> ConvertToConfigExtraList<TKey>(Dictionary<TKey, TValue> dict)
        {
            var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell<TValue>>();

            if (typeof(TKey) == typeof(string))
            {
                foreach (var item in dict)
                {
                    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell<TValue>() { ConfigValue = item.Key as string, FillValue = item.Value });
                }
            }
            else
            {
                foreach (var item in dict)
                {
                    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell<TValue>() { ConfigValue = item.Key.ToString(), FillValue = item.Value });
                }
            }

            return fixedCellsInfoList;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="dt">用来获得列名</param>
        /// <param name="dr">数据源是这个</param>
        public static List<EPPlusConfigSourceFixedCell<TValue>> ConvertToConfigExtraList(DataTable dt, DataRow dr)
        {
            var dict = new Dictionary<string, TValue>();
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                var colName = dt.Columns[i].ColumnName;
                if (!dict.ContainsKey(colName))
                {
                    dict.Add(colName, ValueConvertUtil.ConvertToTValue<TValue>(dr[i]));
                }
                else
                {
                    throw new Exception(nameof(ConvertToConfigExtraList) + "方法异常");
                }
            }
            return ConvertToConfigExtraList<string>(dict);
        }
    }

    public class EPPlusConfigSourceConfigExtras
    {
        public List<EPPlusConfigSourceFixedCell> CellsInfoList { get; set; } = null;

        public EPPlusConfigSourceFixedCell this[string key]
        {
            get
            {
                var cell = GetCellAndTryAdd(key);
                return cell;
            }
            set
            {
                var cell = GetCellAndTryAdd(key);
                cell.FillValue = value;
            }
        }

        private EPPlusConfigSourceFixedCell GetCellAndTryAdd(string key)
        {
            if (string.IsNullOrEmpty(key)) throw new ArgumentNullException($"{nameof(key)}不能为空");
            if (CellsInfoList is null)
            {
                CellsInfoList = new List<EPPlusConfigSourceFixedCell>();
            }

            var cell = CellsInfoList.Find(a => a.ConfigValue == key);
            if (cell is null)
            {
                cell = new EPPlusConfigSourceFixedCell() { ConfigValue = key };
                CellsInfoList.Add(cell);
            }

            return cell;
        }

        public static List<EPPlusConfigSourceFixedCell> ConvertToConfigExtraList<TKey, TValue>(Dictionary<TKey, TValue> dict)
        {
            var fixedCellsInfoList = new List<EPPlusConfigSourceFixedCell>();

#pragma warning disable 184
            if (typeof(TKey) is string)
#pragma warning restore 184
            {
                foreach (var item in dict)
                {
                    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key as string, FillValue = item.Value });
                }
            }
            else
            {
                foreach (var item in dict)
                {
                    fixedCellsInfoList.Add(new EPPlusConfigSourceFixedCell() { ConfigValue = item.Key.ToString(), FillValue = item.Value });
                }
            }

            return fixedCellsInfoList;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="dt">用来获得列名</param>
        /// <param name="dr">数据源是这个</param>
        public static List<EPPlusConfigSourceFixedCell> ConvertToConfigExtraList(DataTable dt, DataRow dr)
        {
            var dict = new Dictionary<string, object>();
            for (int i = 0; i < dr.ItemArray.Length; i++)
            {
                var colName = dt.Columns[i].ColumnName;
                if (!dict.ContainsKey(colName))
                {
                    dict.Add(colName, dr[i] == DBNull.Value ? null : dr[i]);
                }
                else
                {
                    throw new Exception(nameof(ConvertToConfigExtraList) + "方法异常");
                }
            }
            return ConvertToConfigExtraList(dict);
        }
    }

    public class EPPlusConfigSourceHead<TValue> : EPPlusConfigSourceConfigExtras<TValue>
    {
        public static implicit operator EPPlusConfigSourceHead<TValue>(DataTable dt)
        {
            return new EPPlusConfigSourceHead<TValue>()
            {
                CellsInfoList = dt is null || dt.Rows.Count == 0
                    ? new List<EPPlusConfigSourceFixedCell<TValue>>()
                    : EPPlusConfigSourceConfigExtras<TValue>.ConvertToConfigExtraList(dt, dt.Rows[0])
            };
        }

        public static implicit operator EPPlusConfigSourceHead<TValue>(Dictionary<string, TValue> dict) => ConvertToSelf(dict);

        private static EPPlusConfigSourceHead<TValue> ConvertToSelf<TKey>(Dictionary<TKey, TValue> dict)
        {
            return new EPPlusConfigSourceHead<TValue>()
            {
                CellsInfoList = dict is null || dict.Count <= 0
                    ? new List<EPPlusConfigSourceFixedCell<TValue>>()
                    : EPPlusConfigSourceConfigExtras<TValue>.ConvertToConfigExtraList(dict)
            };
        }

        public new object this[string key]
        {
            get => base[key].FillValue;
            set => base[key].FillValue = ValueConvertUtil.ConvertToTValue<TValue>(value);
        }
    }

    public class EPPlusConfigSourceFoot<TValue> : EPPlusConfigSourceConfigExtras<TValue>
    {
        public static implicit operator EPPlusConfigSourceFoot<TValue>(DataTable dt)
        {
            return new EPPlusConfigSourceFoot<TValue>()
            {
                CellsInfoList = dt is null || dt.Rows.Count == 0
                    ? new List<EPPlusConfigSourceFixedCell<TValue>>()
                    : EPPlusConfigSourceConfigExtras<TValue>.ConvertToConfigExtraList(dt, dt.Rows[0])
            };
        }

        public static implicit operator EPPlusConfigSourceFoot<TValue>(Dictionary<string, TValue> dict) => ConvertToSelf(dict);

        private static EPPlusConfigSourceFoot<TValue> ConvertToSelf<TKey>(Dictionary<TKey, TValue> dict)
        {
            return new EPPlusConfigSourceFoot<TValue>()
            {
                CellsInfoList = dict is null || dict.Count <= 0
                    ? new List<EPPlusConfigSourceFixedCell<TValue>>()
                    : EPPlusConfigSourceConfigExtras<TValue>.ConvertToConfigExtraList(dict)
            };
        }

        public new object this[string key]
        {
            get => base[key].FillValue;
            set => base[key].FillValue = ValueConvertUtil.ConvertToTValue<TValue>(value);
        }
    }

    public class EPPlusConfigSourceHead : EPPlusConfigSourceConfigExtras
    {
        public static implicit operator EPPlusConfigSourceHead(DataTable dt)
        {
            return new EPPlusConfigSourceHead()
            {
                CellsInfoList = dt is null || dt.Rows.Count == 0
                    ? new List<EPPlusConfigSourceFixedCell>()
                    : EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dt, dt.Rows[0])
            };
        }

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, string> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, Boolean> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, DateTime> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, sbyte> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, byte> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, UInt16> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, UInt32> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, UInt64> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, Int16> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, Int32> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, Int64> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, float> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, double> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, decimal> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceHead(Dictionary<string, object> dict) => ConvertToSelf(dict);

        private static EPPlusConfigSourceHead ConvertToSelf<TKey, TValue>(Dictionary<TKey, TValue> dict)
        {
            return new EPPlusConfigSourceHead()
            {
                CellsInfoList = dict is null || dict.Count <= 0
                    ? new List<EPPlusConfigSourceFixedCell>()
                    : EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dict)
            };
        }

        public new object this[string key]
        {
            get => base[key].FillValue;
            set
            {
                base[key].FillValue = value;
            }
        }
    }

    public class EPPlusConfigSourceFoot : EPPlusConfigSourceConfigExtras
    {
        public static implicit operator EPPlusConfigSourceFoot(DataTable dt)
        {
            return new EPPlusConfigSourceFoot()
            {
                CellsInfoList = dt is null || dt.Rows.Count == 0
                    ? new List<EPPlusConfigSourceFixedCell>()
                    : EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dt, dt.Rows[0])
            };
        }

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, string> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, Boolean> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, DateTime> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, sbyte> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, byte> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, UInt16> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, UInt32> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, UInt64> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, Int16> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, Int32> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, Int64> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, float> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, double> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, decimal> dict) => ConvertToSelf(dict);

        public static implicit operator EPPlusConfigSourceFoot(Dictionary<string, object> dict) => ConvertToSelf(dict);

        private static EPPlusConfigSourceFoot ConvertToSelf<TKey, TValue>(Dictionary<TKey, TValue> dict)
        {
            return new EPPlusConfigSourceFoot()
            {
                CellsInfoList = dict is null || dict.Count <= 0
                    ? new List<EPPlusConfigSourceFixedCell>()
                    : EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dict)
            };
        }

        public new object this[string key]
        {
            get => base[key].FillValue;
            set => base[key].FillValue = value;
        }
    }

    public class EPPlusConfigSourceBody
    {
        /// <summary>
        /// 所有的配置信息
        /// </summary>
        public List<EPPlusConfigSourceBodyConfig> ConfigList { get; set; } = null;

        /// <summary>
        ///
        /// </summary>
        /// <param name="nth">第几个配置,从1开始</param>
        /// <returns></returns>
        public EPPlusConfigSourceBodyConfig this[int nth]
        {
            get
            {
                if (nth < 1) throw new ArgumentOutOfRangeException($"{nameof(nth)}不能小于1");
                if (ConfigList is null)
                {
                    ConfigList = new List<EPPlusConfigSourceBodyConfig>();
                }

                var bodyConfig = ConfigList.Find(a => a.Nth == nth);
                if (bodyConfig is null)
                {
                    bodyConfig = new EPPlusConfigSourceBodyConfig()
                    {
                        Nth = nth,
                        Option = new EPPlusConfigSourceBodyOption(),
                    };
                    ConfigList.Add(bodyConfig);
                }
                return bodyConfig;
            }
        }
    }

    public class EPPlusConfigSourceBodyConfig
    {
        /// <summary>
        /// 第几个, 从1开始
        /// </summary>
        public int Nth { get; set; }

        /// <summary>
        /// 对应的设置
        /// </summary>
        public EPPlusConfigSourceBodyOption Option = new EPPlusConfigSourceBodyOption();
    }

    public class EPPlusConfigSourceBodyOption
    {
        /// <summary>
        /// 数据源, 对应  EPPlusConfig 的 ConfigLine
        /// </summary>
        public DataTable DataSource { get; set; } = null;

        /// <summary>
        /// 填充方式
        /// </summary>
        public SheetBodyFillDataMethod FillMethod { get; set; } = null;

        /// <summary>
        /// 固定的一些单元格 如表格的汇总栏什么的
        /// </summary>
        public EPPlusConfigSourceConfigExtra ConfigExtra { get; set; } = null;
    }

    public class EPPlusConfigSourceConfigExtra
    {
        /// <summary>
        /// 所有的配置
        /// </summary>
        public List<EPPlusConfigSourceFixedCell> Source { get; set; } = null;

        public static implicit operator EPPlusConfigSourceConfigExtra(DataTable dt)
        {
            return new EPPlusConfigSourceConfigExtra()
            {
                Source = EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dt, dt.Rows[0])
            };
        }

        public static implicit operator EPPlusConfigSourceConfigExtra(Dictionary<string, string> dict)
        {
            return new EPPlusConfigSourceConfigExtra()
            {
                Source = EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dict)
            };
        }

        public static implicit operator EPPlusConfigSourceConfigExtra(Dictionary<string, object> dict)
        {
            return new EPPlusConfigSourceConfigExtra()
            {
                Source = EPPlusConfigSourceConfigExtras.ConvertToConfigExtraList(dict)
            };
        }
    }
}