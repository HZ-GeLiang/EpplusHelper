using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions.Attributes;

namespace EPPlusExtensions.CustomModelType
{
    ///// <summary>
    ///// 为GetList()方法专门设计的,因为GetList中KvSource的方法调用不到,需要通过父类来调用.
    ///// </summary>
    //interface IKVSource
    //{
    //    //bool ContainsKey(object key);
    //    void GetInfoByKey(object key, out bool haveValue, out object value, out bool haveState, out object state);

    //    //ICollection<object> Keys();
    //    //ICollection<object> Values();

    //    //void Clear();
    //    //int Count();
    //}

    public class KvSource<TKey, TValue>/* : IKVSource*/
    {
        private readonly IDictionary<TKey, TValue> _data;
        private readonly IDictionary<TKey, object> _dataState;

        public KvSource()
        {
            this._data = new Dictionary<TKey, TValue>();
            this._dataState = new Dictionary<TKey, object>();
        }

        public KvSource(IDictionary<TKey, TValue> data) => this._data = data;

        public IDictionary<TKey, TValue> Data => this._data;

        //原本是调用 GetState()获得的,新增通过索引获取 
        public IDictionary<TKey, object> DataState => this._dataState;

        public ICollection<TKey> Keys => this._data.Keys;

        public ICollection<TValue> Values => this._data.Values;

        public int Count()
        {
            return this._data.Count;
        }

        public void Add(TKey key, TValue value)
        {
            this._data.Add(key, value);
        }
        public void Add(TKey key, TValue value, object state)
        {
            this.Add(key, value);
            this._dataState.Add(key, state);
        }

        public bool TryAdd(TKey key, TValue value)
        {
            if (this._data.ContainsKey(key))
            {
                return false;
            }
            this._data.Add(key, value);
            return true;
        }
        public bool TryAdd(TKey key, TValue value, object state)
        {
            var result = this.TryAdd(key, value);
            if (result)
            {
                this._dataState.Add(key, state);
            }
            return result;
        }

        public bool TryAdd(object key, object value)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            if (this._data.ContainsKey((TKey)key))
            {
                return false;
            }
            this._data.Add((TKey)key, (TValue)value);
            return true;
        }

        public bool TryAdd(object key, object value, object state)
        {
            var result = this.TryAdd(key, value);
            if (result)
            {
                this._dataState.Add((TKey)key, state);
            }
            return result;
        }

        public KvSource<TKey, TValue> AddRange(IDictionary<TKey, TValue> data)
        {
            foreach (var item in data)
            {
                this._data.Add(item);
            }
            return this;
        }

        public KvSource<TKey, TValue> AddRange(IDictionary<TKey, TValue> data, IDictionary<TKey, object> dataState)
        {
            if (data != null)
            {
                foreach (var item in data)
                {
                    this._data.Add(item);
                }
            }
            if (dataState != null)
            {
                foreach (var item in dataState)
                {
                    this._dataState.Add(item);
                }
            }
            return this;
        }

        public void Clear()
        {
            this._data.Clear();
            this._dataState.Clear();
        }

        public bool ContainsKey(TKey key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            return this._data.ContainsKey(key);
        }

        public bool ContainsKey(object key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            var tkey = (TKey)Convert.ChangeType(key, typeof(TKey));
            return this.ContainsKey(tkey);
        }


        public void GetInfoByKey(object key, out bool haveValue, out object value, out bool haveState, out object state)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            var tkey = (TKey)Convert.ChangeType(key, typeof(TKey));
            haveValue = this._data.ContainsKey(tkey);
            value = haveValue ? this.GetValue(tkey) : (object)default(TValue);
            haveState = this._dataState.ContainsKey(tkey);
            state = haveState ? this.GetState(tkey) : default(object);
        }

        public TValue GetValue(TKey key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            return this._data[key];
        }

        public TValue GetValue(object key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            var tkey = (TKey)Convert.ChangeType(key, typeof(TKey));
            return this.GetValue(tkey);
        }

        public bool HaveState(TKey key, out object state)
        {
            var exists = this._dataState.ContainsKey(key);
            state = exists ? this.GetState(key) : default(object);
            return exists;
        }

        public object GetState(TKey key)
        {
            if (key == null) throw new ArgumentNullException(nameof(key));
            return this._dataState[key];
        }

    }

    /// <summary>
    /// model 的一个类型
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    public class KV<TKey, TValue> : ICustomersModelType
    {
        private TKey _key;
        private TValue _value;
        private object _state;

        public bool HasKey { get; private set; } = false;
        public bool HasValue { get; private set; } = false;
        public bool HasState { get; private set; } = false;

        public KV() { } //这个不能注释.必须提供.存在的理由是方便提供一个默认的对象

        public KV(TKey key)
        {
            this._key = key;
            this.HasKey = true;
        }

        public KV(TKey key, TValue value)
        {
            this._key = key;
            this._value = value;
            this.HasKey = true;
            this.HasValue = true;
        }

        public KV(TKey key, TValue value, object state)
        {
            this._key = key;
            this._value = value;
            this._state = state;
            this.HasKey = true;
            this.HasValue = true;
            this.HasState = true;
        }

        public override string ToString()
        {
            if (this._key == null)
            {
                return "";
            }
            return this._key.ToString();
        }

        /// <summary>
        /// 只读的
        /// </summary>
        public TKey Key => this._key;

        /// <summary>
        /// 只读的
        /// </summary>
        public TValue Value => this._value;

        /// <summary>
        /// 当前key所需要的额外信息;
        /// </summary>
        public object State
        {
            get { return this._state; }
            set { this._state = value; }
        }

        /// <summary>
        /// 保存数据的一个变量
        /// </summary>
        public KvSource<TKey, TValue> KVSource { get; set; }

        //接口部分
        /// <summary>
        /// 有对应的Attribute要搭配使用
        /// </summary>
        public bool HasAttribute { get; set; } = true;

        /// <summary>
        /// HasAttribute == true 才会调用
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="attribute"></param>
        /// <param name="pInfo"></param>
        /// <param name="model"></param>
        /// <param name="value"></param>
        public void RunAttribute<T>(Attribute attribute, PropertyInfo pInfo, T model, string value) where T : class, new()
        {
            if (!(attribute is KVSetAttribute))
            {
                return;
            }
            if (this.KVSource == null)
            {
                throw new ArgumentException($@"检测到KVSetAttribute,但是KVSource却未配置");
            }
            var kvsetAttr = (KVSetAttribute)attribute;
            var haveKvsource = this.KVSource.ContainsKey(value);
            if (kvsetAttr.MustInSet && !haveKvsource)
            {
                var errMsg = !string.IsNullOrEmpty(kvsetAttr.ErrorMessage) && kvsetAttr.ErrorMessage.Length > 0
                      ? EPPlusHelper.FormatAttributeMsg(pInfo.Name, model, value, kvsetAttr.ErrorMessage, kvsetAttr.Args)
                      : $@"属性'{pInfo.Name}'的值:'{value}'未在集合列表中出现.";
                throw new ArgumentException(errMsg, pInfo.Name);
            }

            this.KVSource.GetInfoByKey(value, out bool kv_Value_inKvSource, out object kv_Value, out bool haveState, out object state);

            if (!kv_Value_inKvSource && kvsetAttr.MustInSet)
            {
                var msg = string.IsNullOrEmpty(kvsetAttr.ErrorMessage)
                    ? $@"属性'{pInfo.Name}'的值:'{value}'未在集合中出现."
                    : EPPlusHelper.FormatAttributeMsg(pInfo.Name, model, value, kvsetAttr.ErrorMessage, kvsetAttr.Args);
                throw new ArgumentException(msg, pInfo.Name);
            }

            var typeKVArgs = pInfo.PropertyType.GetGenericArguments();
            var typeKV = typeof(KV<,>).MakeGenericType(typeKVArgs);

            object[] invokeConstructorParameters = new object[]
            {
                typeof(TKey) == typeof(string) ? value : Convert.ChangeType(value,typeof(TKey)),
                kv_Value
            };

            var modelValue = typeKV.GetConstructor(typeKVArgs).Invoke(invokeConstructorParameters);

            if (kv_Value == null) //上面Invoke时, 是调用2个参数的构造方法的,所以,这里要修正HasValue值
            {
                if (!kv_Value_inKvSource)//因为默认值是true,所以,只要修改值为false的情况就可以了
                {
                    typeKV.GetProperty("HasValue").SetValue(modelValue, false);
                }
            }
            if (haveState)
            {
                typeKV.GetProperty("HasState").SetValue(modelValue, true);
                typeKV.GetField("_state", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(modelValue, state);
            }

            pInfo.SetValue(model, modelValue);

        }

        public void SetModelValue<T>(PropertyInfo pInfo, T model, string value) where T : class, new()
        {
            //在RunAttribute已经完成了, 所以,这里是空的
        }

    }

    /// <summary>
    /// 用扩展方法原因: 1.因为当对象为Null时无法获得,2.类型推断
    /// </summary>
    public static class KVExtensionMethod
    {
        public static KvSource<TKey, TValue> CreateKVSource<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>();
        public static Dictionary<TKey, TValue> CreateKVSourceData<TKey, TValue>(this KV<TKey, TValue> source) => new Dictionary<TKey, TValue>();
        //public static Type GetKVSourceType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType();
        public static Type GetKeyType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType().GenericTypeArguments[0];
        public static Type GetValueType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType().GenericTypeArguments[1];
    }

    /// <summary>
    /// 给KvSource搭配使用的.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class KVSetAttribute : Attribute
    {
        /// <summary>
        /// 必须在集合中
        /// </summary>
        public bool MustInSet { get; private set; }

        /// <summary>
        /// 自定义的错误消息
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// 错误消息的参数
        /// </summary>
        public string[] Args { get; private set; }

        public KVSetAttribute() : this(true, null, null) { }


        public KVSetAttribute(bool mustInSet) : this(mustInSet, null, null) { }

        public KVSetAttribute(string errorMessage, params string[] args) : this(true, errorMessage, args) { }

        private KVSetAttribute(bool mustInSet, string errorMessage, params string[] args)
        {
            if (args == null)
            {
                args = new string[0];
            }
            this.MustInSet = mustInSet;
            this.ErrorMessage = errorMessage;
            this.Args = args;
        }

        /// <summary>
        /// Key是属性名字,Value是该属性的类型的 KVSource&lt;TKey,TValue&gt;
        /// </summary>
        public KVSource KVSource = new KVSource();

        public bool AddKVSourceByKey<TKey, TValue>(string key, KvSource<TKey, TValue> value)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentException("key不能为空", nameof(key));
            }
            if (this.KVSource == null)
            {
                return false;
            }
            if (this.KVSource.ContainsKey(key))
            {
                return false;
            }
            this.KVSource.Add(key, value);
            return true;
        }
    }

}
