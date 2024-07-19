namespace EPPlusExtensions.Utils
{
    internal sealed class ValueConvertUtil
    {
        //不支持nullable类型(有需要在替换这个方法)
        public static TValue ConvertToTValue<TValue>(object obj)
        {
            if (obj is TValue)
            {
                return (TValue)obj;
            }

            if (typeof(TValue) == typeof(object))
            {
                return (TValue)obj;
            }

            if (obj is null || obj == DBNull.Value)
            {
                return default;
            }

            if (obj.GetType().IsValueType)
            {
                return (TValue)Convert.ChangeType(obj, typeof(TValue));
            }

            return (TValue)Convert.ChangeType(obj, typeof(TValue));
        }
    }
}
