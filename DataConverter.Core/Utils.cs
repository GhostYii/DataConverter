using System.Text;

namespace DataConverter.Core
{
    internal static class Utils
    {
        private static char[] InvaildChars = new char[]
        {
            '~', '!', '@', '#', '$', '%', '^', '&', '*',
            '(', ')', '-', '=', '+', '[', '{', ']', '}',
            ';', ':', '\'', '\"', ',', '<', '.', '>', '/', '?'
        };

        private static string[] InvaliedFieldNames = new string[]
        {
            "abstract", "event", "new", "struct", "as", "explicit", "null", "switch", "base", "extern",
            "this", "false", "operator", "throw", "break", "finally", "out", "true",
            "fixed", "override", "try", "case", "params", "typeof", "catch", "for",
            "private", "foreach", "protected", "checked", "goto", "public",
            "unchecked", "class", "if", "readonly", "unsafe", "const", "implicit", "ref",
            "continue", "in", "return", "using", "virtual", "default",
            "interface", "sealed", "volatile", "delegate", "internal", "do", "is",
            "sizeof", "while", "lock", "stackalloc", "else", "static", "enum",
            "namespace",
            "object", "bool", "byte", "float", "uint", "char", "ulong", "ushort",
            "decimal", "int", "sbyte", "short", "double", "long", "string", "void",
            "partial",  "yield",  "where"
        };

        // 字符串转变量名，失败返回null
        public static string ToFieldName(string name)
        {
            name = name.Trim().TrimStart('#').TrimEnd('*');

            // 空名字
            if (string.IsNullOrEmpty(name))
                return null;
            
            // 含有非法字符
            if (name.Any((ch) => InvaildChars.Contains(ch)))
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

        public static bool IsValidIdentifier(string value)
        {
            return !InvaliedFieldNames.Contains(value);
        }
    }
}
