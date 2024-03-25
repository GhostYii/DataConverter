using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Media;

namespace DataConverter
{
    internal static class Commands
    {
        private class CommandInfo
        {
            public object obj;
            public MethodInfo method;
        }

        public static TerminalControl Terminal { get; set; }

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
                Terminal?.AppendRaw($"未注册的命令\"{name}\"", TerminalControl.LogType.Error.ToBrush());
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
                catch (Exception e) { Terminal?.AppendRaw(e.Message, TerminalControl.LogType.Error.ToBrush()); }

                return;

            }

            TipMethodParams(name);
        }
        public static void Execute(string name, params object[] args)
        {
            if (!_cmds.ContainsKey(name.ToLower()))
            {
                Terminal?.AppendRaw($"未注册的命令\"{name}\"", TerminalControl.LogType.Error.ToBrush());
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
                catch (Exception e) { Terminal?.AppendRaw(e.Message, TerminalControl.LogType.Error.ToBrush()); }

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
            sb.AppendLine($"命令\"{name}\"参数不匹配");
            foreach (var cmd in _cmds[name])
            {
                sb.AppendFormat("- {0}(", name);

                for (int i = 0; i < cmd.method.GetParameters().Length; i++)
                {
                    Type type = cmd.method.GetParameters()[i].ParameterType;
                    sb.Append(type);

                    if (i != cmd.method.GetParameters().Length - 1)
                        sb.Append(", ");
                }
                sb.Append(")\n");
            }

            Terminal?.AppendRaw(sb.ToString(), TerminalControl.LogType.Warning.ToBrush());
        }

        #endregion

        #region build-in commands

        [CMD("help", "显示帮助信息"), CMD("man", "显示帮助信息")]
        private static void Help(string cmdName)
        {
            if (!_cmds.ContainsKey(cmdName.ToLower()))
            {
                Terminal?.AppendRaw($"未注册的命令\"{cmdName}\"", TerminalControl.LogType.Error.ToBrush());
                return;
            }

            cmdName = cmdName.Trim().ToLower();
            StringBuilder sb = new StringBuilder();

            foreach (var cmd in _cmds[cmdName])
            {
                sb.AppendFormat("{0}(", cmdName);
                bool isFirst = true;
                foreach (var parameter in cmd.method.GetParameters())
                {
                    sb.Append(parameter.ParameterType);
                    if (!isFirst)
                    {
                        sb.Append(",");
                        isFirst = false;
                    }
                }
                sb.AppendFormat("):\n{0}\n", cmd.method.GetCustomAttributes<CMDAttribute>().ToArray()[0].Desc);
            }

            Terminal?.AppendRaw(sb.ToString(), Brushes.LightGreen);
        }

        [CMD("cls", "清屏"), CMD("clear", "清屏")]
        private static void Cls()
        {
            Terminal?.Clear();
        }

        #endregion        
    }
}
