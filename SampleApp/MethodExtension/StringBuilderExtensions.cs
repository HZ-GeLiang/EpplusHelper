using System.Text;

namespace SampleApp.MethodExtension
{
    internal static class StringBuilderExtensions
    {
        public static StringBuilder RemoveLastChar(this StringBuilder value)
        {
            return value == null || value.Length <= 0 ? value : value.Remove(value.Length - 1, 1);
        }
        public static StringBuilder RemoveLastChar(this StringBuilder value, int count)
        {
            if (value == null) throw new System.ArgumentNullException(nameof(value));
            if (count <= 0) throw new System.ArgumentException(nameof(count));
            if (count > value.Length) throw new System.ArgumentException(nameof(count));
            return value.Remove(value.Length - count, count);
        }

        public static StringBuilder RemoveLastChar(this StringBuilder value, char c)
        {
            if (value == null) throw new System.ArgumentNullException(nameof(value));
            return value.Length <= 0 ? value : value[value.Length - 1] == c ? value.RemoveLastChar() : value;
        }
        public static StringBuilder RemoveLastChar(this StringBuilder value, string str)
        {
            if (value == null) throw new System.ArgumentNullException(nameof(value));
            if (str == null || str.Length <= 0)
            {
                return value;
            }
            if (str.Length > value.Length)
            {
                //throw new System.ArgumentException($"{nameof(str)}的长度大于{nameof(value)}");
                return value;
            }

            int eachCount = 0;
            for (int i = str.Length; i > 0; i--)
            {
                eachCount++;
                if (str[i - 1] != value[value.Length - eachCount])
                {
                    return value;
                }
            }
            return value.Replace(str, string.Empty, value.Length - str.Length, str.Length);

        }


    }
}
