﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions
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

    }
}
