﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConverter.Core
{
    internal static class Extension
    {
        public static bool IsList(this Type type)
        {
            return type != null && type.IsGenericType && type.GetGenericTypeDefinition().IsAssignableFrom(typeof(List<>));
        }

        public static bool IsDictionary(this Type type)
        {
            return type != null && type.IsGenericType && type.GetGenericTypeDefinition().IsAssignableFrom(typeof(Dictionary<,>));
        }

        /// <summary>
        /// is value type (string is TRUE)
        /// </summary>
        public static bool IsValueType(this CellValueType val)
        {
            return val == CellValueType.Int || val == CellValueType.Float || val == CellValueType.String;
        }

    }
}
