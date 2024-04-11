using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DataConverter.Core
{
    [AttributeUsage(AttributeTargets.Method)]
    internal class ExcelTypeParserAttribute : Attribute
    {
        public string[] Types { get; private set; }

        public ExcelTypeParserAttribute(string type)
        {
            Types = new string[] { type };
        }

        public ExcelTypeParserAttribute(params string[] types)
        {
            if (types.Length < 1)
                throw new ArgumentException("至少定义一种类型");

            HashSet<string> hash = new HashSet<string>(types);
            Types = hash.ToArray();
        }

        public bool CheckValidMethod(MethodInfo method)
        {
            if (method.ReturnParameter.ParameterType != typeof(CellType))
                return false;

            var args = method.GetParameters();
            if (args.Length != 2)
                return false;

            if (args[0].ParameterType != typeof(string) &&
                args[1].ParameterType != typeof(string))
                return false;

            return true;
        }
    }
}
