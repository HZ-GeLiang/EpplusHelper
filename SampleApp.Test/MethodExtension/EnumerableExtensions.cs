using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApp.Test
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> GetEmpty<T>(this IEnumerable<T> enumerable)
        {
            return Enumerable.Empty<T>();
        }
    }
}
