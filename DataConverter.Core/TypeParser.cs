using System.Reflection;
using System.Text;

namespace DataConverter.Core
{
    public static class TypeParser
    {
        private static Dictionary<string, MethodInfo> _typeParsers = new Dictionary<string, MethodInfo>();

        static TypeParser() { LoadParser(typeof(TypeParser));  }

        public static void LoadParser(Type type)
        {
            _typeParsers.Clear();

            var methods = type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            foreach (var method in methods)
            {
                var attrib = method.GetCustomAttribute<ExcelTypeParserAttribute>();
                if (attrib == null)
                    continue;

                if (!attrib.CheckValidMethod(method))
                {
                    Console.PrintError($"类型解析方法（{method.Name}）定义不正确，原型必须为CellType MethodName(string type, string subType)");
                    continue;
                }

                foreach (var t in attrib.Types)
                {
                    if (!_typeParsers.ContainsKey(t))
                    {
                        _typeParsers[t] = method;
                    }
                    else
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendFormat("{0} {1}(", method.ReturnType, method.Name);
                        var args = method.GetParameters();
                        for (int i = 0; i < args.Length; ++i)
                        {
                            var arg = args[i];
                            sb.Append(arg.ParameterType);
                            if (i != args.Length - 1)
                                sb.Append(',');
                        }
                        sb.Append(')');

                        Console.PrintError($"重复的{type}类型解析方法，方法\"{sb}\"将被舍弃");
                        continue;
                    }
                }
            }

            Console.Print($"解析函数加载完成，成功加载{_typeParsers.Count}个函数");
        }

        internal static CellType Parse(string typeArg)
        {            
            string[] typeStr = SplitType(typeArg);
            string keyName = typeStr[0].Trim().ToLower();
            if (!_typeParsers.ContainsKey(keyName))
                return null;

            return _typeParsers[keyName].Invoke(null, new object[] { typeStr[0], typeStr[1] }) as CellType;
        }

        internal static string[] SplitType(string typeArg)
        {
            StringBuilder sb = new StringBuilder();
            int offset = 0;
            foreach (char ch in typeArg.Trim().ToLower())
            {
                if (ch != Const.TYPE_SPLIT_CHAR)
                {
                    sb.Append(ch);
                    ++offset;
                }
                else
                {
                    break;
                }
            }

            return new string[2] { sb.ToString(), offset + 1 >= typeArg.Length ? string.Empty : typeArg.Substring(offset + 1) };
        }

        // TODO: support auto
        //[ExcelTypeParser("auto")]
        //private static CellType AutoTypeParser(string type, string define)
        //{
        //    return CellType.Default;
        //}

        [ExcelTypeParser("int")]
        private static CellType IntParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Int };
        }

        [ExcelTypeParser("float")]
        private static CellType FloatParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Float };
        }

        [ExcelTypeParser("bool", "boolean")]
        private static CellType BoolParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Bool };
        }

        [ExcelTypeParser("string", "str")]
        private static CellType StringParser(string type, string subType)
        {
            return new CellType { type = CellValueType.String };
        }

        [ExcelTypeParser("object", "obj")]
        private static CellType ObjectParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Object, objName = subType };
        }

        [ExcelTypeParser("array", "arr", "list")]
        private static CellType ArrayParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Array, subType = Parse(subType) };
        }

        [ExcelTypeParser("map", "dict", "pairs", "hash")]
        private static CellType DictionaryParser(string type, string subType)
        {
            return new CellType { type = CellValueType.Map, subType = Parse(subType) };
        }

    }
}
