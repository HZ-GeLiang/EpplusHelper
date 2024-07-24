using EPPlusExtensions.CustomModelType;

namespace SampleApp
{
    internal class Helper
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
            if (a is null && b is null)
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