﻿using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace DataConverter
{
    internal static class TypeParser
    {
        private static Dictionary<string, MethodInfo> _typeParsers = new Dictionary<string, MethodInfo>();

        static TypeParser()
        {
            var methods = typeof(TypeParser).GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
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

                foreach (var type in attrib.Types)
                {
                    if (!_typeParsers.ContainsKey(type))
                    {
                        _typeParsers[type] = method;
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
        }

        public static CellType Parse(string typeArg)
        {
            //Console.Print($"解析类型：{typeArg}");
            string[] typeStr = SplitType(typeArg);
            string keyName = typeStr[0].Trim().ToLower();
            if (!_typeParsers.ContainsKey(keyName))
                return null;

            return _typeParsers[keyName].Invoke(null, new object[] { typeStr[0], typeStr[1] }) as CellType;
        }

        public static string[] SplitType(string typeArg)
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


        //[ExcelTypeParser("auto")]
        //private static CellType AutoTypeParser(string type, string define)
        //{
        //    return CellType.Default;
        //}

        [ExcelTypeParser("int")]
        private static CellType IntParser(string type, string subType)
        {
            return new CellType { type = ValueType.Int };
        }

        [ExcelTypeParser("float")]
        private static CellType FloatParser(string type, string subType)
        {
            return new CellType { type = ValueType.Float };
        }

        [ExcelTypeParser("string", "str")]
        private static CellType StringParser(string type, string subType)
        {
            return new CellType { type = ValueType.String };
        }

        [ExcelTypeParser("object", "obj")]
        private static CellType ObjectParser(string type, string subType)
        {
            return new CellType { type = ValueType.Object, objName = subType };
        }

        [ExcelTypeParser("array", "arr", "list")]
        private static CellType ArrayParser(string type, string subType)
        {
            return new CellType { type = ValueType.Array, subType = Parse(subType) };
        }

        [ExcelTypeParser("map", "dict", "pairs")]
        private static CellType DictionaryParser(string type, string subType)
        {
            return new CellType { type = ValueType.Map, subType = Parse(subType) };
        }

    }
}
