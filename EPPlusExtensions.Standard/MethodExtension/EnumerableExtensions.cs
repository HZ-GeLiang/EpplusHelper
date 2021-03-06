﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions
{
    internal static class EnumerableExtensions
    {
        public static IEnumerable<TSource> GetRepeatBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
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
