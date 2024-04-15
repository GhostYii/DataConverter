using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConverter.Core
{
    // apis
    public static class DC
    {
        internal static Dictionary<string, Type> _extendTypes = new Dictionary<string, Type>();

        internal static Type GetType(string name)
        {
            if(!_extendTypes.ContainsKey(name))
                return null;

            return _extendTypes[name];
        }

        public static void RegisterTypes(params Type[] types)
        {
            foreach (var type in types)
            {
                if(_extendTypes.ContainsKey(type.Name))
                {
                    Console.PrintError($"Type名{type.Name}已被{_extendTypes[type.Name].FullName}注册");
                    continue;
                }

                _extendTypes[type.Name] = type;
            }
        }

        public static void RemoveExtendType(Type type)
        {
            _extendTypes.Remove(type.Name);
        }

    }
}
