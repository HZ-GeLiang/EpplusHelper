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

        public KVSetAttribute(string name) => InitConstructor(name, true, null, new string[0]);

        public KVSetAttribute(string name, bool mustInSet) => InitConstructor(name, mustInSet, null, new string[0]);

        public KVSetAttribute(string name, string errorMessage, params string[] args) => InitConstructor(name, true, errorMessage, args);
        public KVSetAttribute(string name, bool mustInSet, string errorMessage, params string[] args) => InitConstructor(name, mustInSet, errorMessage, args);

        private void InitConstructor(string name, bool mustInSet, string errorMessage, string[] args)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name)); 
            this.Name = name;
            this.MustInSet = mustInSet;
            this.ErrorMessage = errorMessage;
            this.Args = args;
        }
    }

    interface IKVSource
    {
        bool ContainsKey(object key, out object value);

        //ICollection<object> Keys();
        //ICollection<object> Values();
        //bool ContainsKey(object key);
        //void Clear();
        //int Count();
        //object GetValue();
    }

    public class KvSource<TKey, TValue> : IKVSource
    {
        private IDictionary<TKey, TValue> _data;

        public KvSource() => this._data = new Dictionary<TKey, TValue>();

        public KvSource(IDictionary<TKey, TValue> data) => this._data = data;

        public bool TryAdd(TKey key, TValue value)
        {
            if (this._data.ContainsKey(key))
            {
                return false;
            }
            else
            {
                this._data.Add(key, value);
                return true;
            }
        }

        public bool TryAdd(object key, object value)
        {
            if (key == null) throw new Exception("Key为null");
            if (this._data.ContainsKey((TKey)key))
            {
                return false;
            }
            else
            {
                this._data.Add((TKey)key, (TValue)value);
                return true;
            }
        }

        public KvSource<TKey, TValue> AddRange(IDictionary<TKey, TValue> data)
        {
            foreach (var item in data)
            {
                this._data.Add(item);
            }
            return this;
        }

        public void Clear()
        {
            this._data.Clear();
        }

        public bool ContainsKey(object key, out object value)
        {
            if (key == null) throw new Exception("Key为null");
            var _key = (TKey)Convert.ChangeType(key, typeof(TKey));
            var isExists = this._data.ContainsKey(_key);
            value = isExists ? this._data[_key] : (object)default(TValue);
            return isExists;
        }

        public bool ContainsKey(object key)
        {
            if (key == null) throw new Exception("Key为null");
            var _key = (TKey)Convert.ChangeType(key, typeof(TKey));
            var isExists = this._data.ContainsKey(_key);
            return isExists;
        }

        public int Count => this._data.Count;
        public ICollection<TKey> Keys => this._data.Keys;
        public ICollection<TValue> Values => this._data.Values;
        public IDictionary<TKey, TValue> Data => this._data;

    }

    public class KV<TKey, TValue>
    {
        private KeyValuePair<TKey, TValue> _kv;

        public bool HasValue { get; set; } = false;

        public KV() { } //这个不能注释.必须提供.存在的理由是方便提供一个默认的对象

        public KV(TKey key, TValue value) => _kv = new KeyValuePair<TKey, TValue>(key, value);

        public override string ToString() => this._kv.Key.ToString();

        public TKey Key => this._kv.Key;
        public TValue Value => this._kv.Value;

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