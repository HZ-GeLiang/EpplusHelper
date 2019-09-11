using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Helper
{
    internal class Utils
    {
        public static TValue ConvertToTValue<TValue>(object obj)
        {
            if (obj is TValue)
            {
                return (TValue)obj;
            }

            if (typeof(TValue) == typeof(Object))
            {
                return (TValue)obj;
            }

            if (obj == null || obj == DBNull.Value)
            {
                return default(TValue);
            }

            if (obj.GetType().IsValueType)
            {
                return (TValue)Convert.ChangeType(obj, typeof(TValue));
            }

            return (TValue)Convert.ChangeType(obj, typeof(TValue));

        }
    }
}
