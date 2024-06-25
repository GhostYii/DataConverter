using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace DataConverter.CLI
{
    internal static class Commands
    {
        private class CommandInfo
        {
            public object obj;
            public MethodInfo method;
        }

        private static Dictionary<string, HashSet<CommandInfo>> _cmds = new Dictionary<string, HashSet<CommandInfo>>();

        static Commands()
        {
            RegisterAllCommandByType(typeof(Commands));
        }

        #region Framework

        public static void InitialCommandByType(params Type[] type)
        {
            foreach (var t in type)
            {
                RegisterAllCommandByType(t);
            }
        }
        public static void InitialCommandByObject<T>(params T[] obj)
        {
            foreach (var o in obj)
            {
                RegisterAllCommandByObject(o);
            }
        }

        public static void RegisterAllCommandByType(Type type)
        {
            var methods = type.GetMethods
                          (
                              BindingFlags.Static |
                              BindingFlags.Public |
                              BindingFlags.NonPublic
                          );
            foreach (var method in methods)
            {
                foreach (CMDAttribute cmdAttr in method.GetCustomAttributes<CMDAttribute>())
                {
                    if (string.IsNullOrEmpty(cmdAttr.Name))
                        RegisterCommand(method.Name.Trim().ToLower(), null, method);
                    else
                        RegisterCommand(cmdAttr.Name.Trim().ToLower(), null, method);
                }
            }
        }
        public static void RegisterAllCommandByObject(object obj)
        {
            if (obj == null)
                return;

            var methods = obj.GetType().GetMethods
                          (
                              BindingFlags.Instance |
                              BindingFlags.InvokeMethod |
                              BindingFlags.Static |
                              BindingFlags.Public |
                              BindingFlags.NonPublic
                          );
            foreach (var method in methods)
            {
                foreach (CMDAttribute cmdAttr in method.GetCustomAttributes<CMDAttribute>())
                {
                    if (string.IsNullOrEmpty(cmdAttr.Name))
                        RegisterCommand(method.Name.Trim().ToLower(), obj, method);
                    else
                        RegisterCommand(cmdAttr.Name.Trim().ToLower(), obj, method);
                }
            }
        }

        public static void Execute(string name, string[] args)
        {
            if (!_cmds.ContainsKey(name.ToLower()))
            {                
                Console.WriteLine($"unregistered command \"{name}\"");
                return;
            }

            name = name.Trim().ToLower();

            foreach (var cmd in _cmds[name])
            {
                var parameters = cmd.method.GetParameters();
                if (parameters.Length != args.Length)
                    continue;

                bool isMatch = true;
                object[] argsWithType = new object[args.Length];

                // support int/float/string type parameters
                for (int i = 0; i < parameters.Length; i++)
                {
                    if (parameters[i].ParameterType == typeof(int) && int.TryParse(args[i], out int intRes))
                        argsWithType[i] = intRes;
                    else if (parameters[i].ParameterType == typeof(float) && float.TryParse(args[i], out float floatRes))
                        argsWithType[i] = floatRes;
                    else
                        argsWithType[i] = args[i];

                    if (parameters[i].ParameterType != argsWithType[i].GetType())
                    {
                        isMatch = false;
                        break;
                    }
                }

                if (!isMatch)
                    continue;

                try { cmd.method.Invoke(cmd.obj, argsWithType); }
                catch (Exception e) { Console.WriteLine(e.Message); }

                return;

            }

            TipMethodParams(name);
        }
        public static void Execute(string name, params object[] args)
        {
            if (!_cmds.ContainsKey(name.ToLower()))
            {                
                Console.WriteLine($"unregistered command \"{name}\"");
                return;
            }

            name = name.Trim().ToLower();

            foreach (var cmd in _cmds[name])
            {
                var parameters = cmd.method.GetParameters();
                if (parameters.Length != args.Length)
                    continue;

                bool isMatch = true;

                for (int i = 0; i < parameters.Length; i++)
                {
                    if (parameters[i].ParameterType != args[i].GetType())
                    {
                        isMatch = false;
                        break;
                    }
                }

                if (!isMatch)
                    continue;

                try { cmd.method.Invoke(cmd.obj, args); }
                catch (Exception e) { Console.WriteLine(e.Message); }

                return;
            }

            TipMethodParams(name);
        }

        public static void RegisterCommand(string name, object invokeObj, MethodInfo method)
        {
            if (string.IsNullOrEmpty(name))
                return;

            if (!_cmds.ContainsKey(name))
                _cmds[name] = new HashSet<CommandInfo>();

            _cmds[name].Add(new CommandInfo() { obj = invokeObj, method = method });
            //StringBuilder sb = new StringBuilder();
            //sb.AppendFormat("已注册命令{0}(", name);
            //foreach (var parameter in method.GetParameters())
            //{
            //    sb.AppendFormat("{0},", parameter.ParameterType);
            //}
            //sb.Append(")");
            //Terminal?.AppendRaw(sb.ToString());
        }

        private static void TipMethodParams(string name)
        {
            if (!_cmds.ContainsKey(name))
                return;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"command \"{name}\" arguments mismatch.");
            foreach (var cmd in _cmds[name])
            {
                sb.AppendFormat("- {0}(", name);

                for (int i = 0; i < cmd.method.GetParameters().Length; i++)
                {
                    var paramInfo = cmd.method.GetParameters()[i];
                    sb.AppendFormat("{0} {1}", paramInfo.ParameterType.Name, paramInfo.Name);

                    if (i != cmd.method.GetParameters().Length - 1)
                        sb.Append(", ");
                }
                sb.Append(")\n");
            }
            
            Console.WriteLine(sb.ToString());
        }

        #endregion

        #region build-in commands

        [CMD("all", "show all cmds")]
        private static void All()
        {
            foreach (var (name, cmds) in _cmds)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"command \"{name}\"");
                foreach (var cmd in cmds)
                {
                    sb.AppendFormat("- {0}(", name);

                    for (int i = 0; i < cmd.method.GetParameters().Length; i++)
                    {
                        var paramInfo = cmd.method.GetParameters()[i];
                        sb.AppendFormat("{0} {1}", paramInfo.ParameterType.Name, paramInfo.Name);

                        if (i != cmd.method.GetParameters().Length - 1)
                            sb.Append(", ");
                    }
                    sb.Append(")\n");
                }
                Console.WriteLine(sb.ToString());
            }
        }

        [CMD("help", "show tip message"), CMD("man", "show tip message")]
        private static void Help(string cmdName)
        {
            if (!_cmds.ContainsKey(cmdName.ToLower()))
            {                
                Console.WriteLine($"unregisterd command \"{cmdName}\"");
                return;
            }

            cmdName = cmdName.Trim().ToLower();
            StringBuilder sb = new StringBuilder();

            foreach (var cmd in _cmds[cmdName])
            {
                sb.AppendFormat("{0}(", cmdName);
                for (int i = 0; i < cmd.method.GetParameters().Length; i++)
                {
                    var paramInfo = cmd.method.GetParameters()[i];
                    sb.AppendFormat("{0} {1}", paramInfo.ParameterType.Name, paramInfo.Name);

                    if (i != cmd.method.GetParameters().Length - 1)
                        sb.Append(", ");
                }
                sb.AppendFormat("):\n{0}\n", cmd.method.GetCustomAttributes<CMDAttribute>().ToArray()[0].Desc);
            }
            
            Console.WriteLine(sb.ToString());
        }

        [CMD("cls", "clear screen"), CMD("clear", "clear screen")]
        private static void Cls()
        {
            Console.Clear();
        }

        #endregion        
    }
}
