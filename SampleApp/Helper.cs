using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlusExtensions.Attributes;
using EPPlusExtensions.CustomModelType;

namespace SampleApp
{
    class Helper
    {
        internal static int GetHashCode_KV<TKey, TValue>(KV<TKey, TValue> kv)
        {
            return kv.Key.GetHashCode() +
                   kv.Value.GetHashCode() +
                   kv.State.GetHashCode() +
                   kv.HasKey.GetHashCode() +
                   kv.HasValue.GetHashCode() +
                   kv.HasState.GetHashCode();
        }

        internal static bool GetEquals_KV<TKey, TValue>(KV<TKey, TValue> a, KV<TKey, TValue> b)
        {
            if (a == null && b == null)
            {
                return true;
            }
            return a?.Key?.GetHashCode() == b?.Key?.GetHashCode() &&
                   a?.Value?.GetHashCode() == b?.Value?.GetHashCode() &&
                   a?.State?.GetHashCode() == b?.State?.GetHashCode() &&
                   a.HasKey.GetHashCode() == b.HasKey.GetHashCode() &&
                   a.HasValue.GetHashCode() == b.HasValue.GetHashCode() &&
                   a.HasState.GetHashCode() == b.HasState.GetHashCode();
        }
    }
}
