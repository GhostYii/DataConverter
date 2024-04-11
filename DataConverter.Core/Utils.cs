using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConverter.Core
{
    internal static class Utils
    {
        // 字符串转变量名，失败返回null
        public static string ToFieldName(string name)
        {
            name = name.Trim().TrimStart('#').TrimEnd('*');

            // 空名字
            if (string.IsNullOrEmpty(name))
                return null;

            char[] invaildChars = new char[]
            {
                '~', '!', '@', '#', '$', '%', '^', '&', '*',
                '(', ')', '-', '=', '+', '[', '{', ']', '}',
                ';', ':', '\'', '\"', ',', '<', '.', '>', '/', '?'
            };

            // 含有非法字符
            if (name.Any((ch) => invaildChars.Contains(ch)))
                return null;

            StringBuilder sb = new StringBuilder();
            string[] splits = name.Split('_');

            // 数字开头加上下划线前缀            
            if (int.TryParse(name.Substring(0, 1), out int intRes))
                sb.Append('_');

            if (splits.Length > 0)
            {
                foreach (var item in splits)
                {
                    sb.Append(item.Substring(0, 1).ToUpper());
                    sb.Append(item.Substring(1, item.Length - 1));
                }
            }

            return sb.ToString();
        }

        // 自动分析类型
        public static ValueType TryParseValueType(object value)
        {
            string strValue = value.ToString();
            if (int.TryParse(strValue, out int intRes))
                return ValueType.Int;

            if (float.TryParse(strValue, out float floatRes))
                return ValueType.Float;

            return ValueType.String;
        }
    }
}
