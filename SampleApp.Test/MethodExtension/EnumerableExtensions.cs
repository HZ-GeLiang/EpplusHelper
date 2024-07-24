using System.Collections.Generic;
using System.Linq;

namespace SampleApp.Test
{
    internal static class EnumerableExtensions
    {
        public static IEnumerable<T> GetEmpty<T>(this IEnumerable<T> enumerable)
        {
            return Enumerable.Empty<T>();
        }
    }
}