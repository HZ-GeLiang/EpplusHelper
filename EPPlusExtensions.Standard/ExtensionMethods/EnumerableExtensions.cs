namespace EPPlusExtensions.ExtensionMethods
{
    internal static class EnumerableExtensions
    {
        public static IEnumerable<TSource> GetRepeat<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            if (source is null)
            {
                throw new ArgumentException($@"{nameof(source)} can not be null");
            }

            var hashSet = new HashSet<TKey>();
            foreach (var item in source)
            {
                if (!hashSet.Add(keySelector(item)))
                {
                    yield return item;
                }
            }
        }
    }
}