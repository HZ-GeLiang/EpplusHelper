using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Attributes
{

    public class KVSetAttribute : System.Attribute
    {
        /// <summary>
        /// 必须在集合中
        /// </summary>
        public bool MustInSet { get; private set; }

        /// <summary>
        ///  集合名字
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 自定义的错误消息
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// 错误消息的参数
        /// </summary>
        public string[] Args { get; private set; }


        public KVSetAttribute(string name)
        {
            this.Name = name;
            this.MustInSet = true;
            this.Args = new string[0];
        }
        public KVSetAttribute(string name, bool mustInSet)
        {
            this.Name = name;
            this.MustInSet = mustInSet;
            this.Args = new string[0];
        }

        public KVSetAttribute(string name, bool mustInSet, string errorMessage, params string[] args)
        {
            this.Name = name;
            this.MustInSet = mustInSet;
            this.ErrorMessage = errorMessage;
            this.Args = args;
        }

        public KVSetAttribute(string name, string errorMessage, params string[] args)
        {
            this.Name = name;
            this.MustInSet = true;
            this.ErrorMessage = errorMessage;
            this.Args = args;
        }
    }

    public abstract class KVSource
    {
        public abstract bool ContainsKey(object key, out object value);
        //public abstract object GetValueByKey(object key);
    }

    public class KvSource<TKey, TValue> : KVSource
    {
        private IDictionary<TKey, TValue> _data;

        public KvSource() => this._data = new Dictionary<TKey, TValue>();

        public KvSource(IDictionary<TKey, TValue> data) => this._data = data;

        #region 注释的代码
        //public ICollection<TKey> Keys => this._data.Keys;
        //public ICollection<TValue> Values => this._data.Values;
        //public bool ContainsKey(TKey key) => this._data.ContainsKey(key);

        //public bool ContainsKey(object key)
        //{
        //    if (key == null) throw new Exception("Key为null");
        //    return key.GetType() == typeof(TKey)
        //        ? ContainsKey((TKey)key)
        //        : this._data.ContainsKey((TKey)Convert.ChangeType(key, typeof(TKey)));
        //}

        //public void Add(TKey key, TValue value) => this._data.Add(key, value);

        public KvSource<TKey, TValue> AddRange(IDictionary<TKey, TValue> data)
        {
            foreach (var item in data)
            {
                this._data.Add(item);
            }
            return this;
        }

        //public bool Remove(TKey key) => this._data.Remove(key);
        //public bool TryGetValue(TKey key, out TValue value) => this._data.TryGetValue(key, out value);

        //public void Clear() => this._data.Clear();

        //public int Count => this._data.Count;

        //public TValue this[TKey key] => this._data[key];

        //public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator() => this._data.GetEnumerator();

        #endregion

        public override bool ContainsKey(object key, out object value)
        {
            if (key == null) throw new Exception("Key为null");
            var _key = (TKey)Convert.ChangeType(key, typeof(TKey));
            var isExists = this._data.ContainsKey(_key);
            value = isExists ? this._data[_key] : (object)default(TValue);
            return isExists;
        }
    }

    public class KV<TKey, TValue>
    {
        private KeyValuePair<TKey, TValue> _kv;
        private Type TKeyType;
        private Type TValueType;

        public bool HasValue { get; set; } = false;

        public KV()
        {
            this.TKeyType = typeof(TKey);
            this.TValueType = typeof(TValue);
        }

        public KV(TKey key, TValue value)
        {
            _kv = new KeyValuePair<TKey, TValue>(key, value);
            this.TKeyType = typeof(TKey);
            this.TValueType = typeof(TValue);
        }

        public TKey Key => this._kv.Key;

        public TValue Value => this._kv.Value;
        public Type KeyType => this.TKeyType;

        public Type ValueType => this.TValueType;

        public override string ToString() => this._kv.ToString();

        //public KVSource<TKey, TValue> CreateKVSource() =>  return new KVSource<TKey, TValue>();//这样写,对象必须不能为空,改用扩展方法
    }

    public static class KVExtensionMethod
    {
        public static KvSource<TKey, TValue> CreateKVSource<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>();
        public static Type GetKVSourceType<TKey, TValue>(this KV<TKey, TValue> source) => new KvSource<TKey, TValue>().GetType();

    }
}