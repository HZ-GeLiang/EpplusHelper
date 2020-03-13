using System;
using System.Collections.Generic;

namespace EPPlusExtensions.Attributes
{
    /// <summary>
    /// 给KvSource搭配使用的.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class KVSetAttribute : Attribute
    {
        /// <summary>
        ///  集合名字
        /// </summary>
        public string Name { get; private set; }

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

        public KVSetAttribute(string name)
            : this(name, true, null, new string[0])
        { }

        public KVSetAttribute(string name, bool mustInSet)
            : this(name, mustInSet, null, new string[0])
        { }

        public KVSetAttribute(string name, string errorMessage, params string[] args)
            : this(name, true, errorMessage, args)
        { }

        public KVSetAttribute(string name, bool mustInSet, string errorMessage, params string[] args)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));
            this.Name = name;
            this.MustInSet = mustInSet;
            this.ErrorMessage = errorMessage;
            this.Args = args;
        }
    }

    /// <summary>
    /// 为GetList()方法专门设计的,因为GetList中KvSource的方法调用不到,需要通过父类来调用.
    /// </summary>
    interface IKVSource
    {
        //bool ContainsKey(object key);
        void GetInfoByKey(object key, out bool haveValue, out object value, out bool haveState, out object state);

        //ICollection<object> Keys();
        //ICollection<object> Values();

        //void Clear();
        //int Count();
    }

    public class KvSource<TKey, TValue> : IKVSource
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

    public class KV<TKey, TValue>
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

    }

    /// <summary>
    /// 因为当对象为Null时,无法获得, 所以,改用扩展方法
    /// </summary>
    public static class KVExtensionMethod
    {
        public static KvSource<TKey, TValue> CreateKVSource<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>();
        public static Dictionary<TKey, TValue> CreateKVSourceData<TKey, TValue>(this KV<TKey, TValue> source) => new Dictionary<TKey, TValue>();
        //public static Type GetKVSourceType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType();
        public static Type GetKeyType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType().GenericTypeArguments[0];
        public static Type GetValueType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType().GenericTypeArguments[1];

    }
}