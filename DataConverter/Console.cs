using System.Windows.Media;

namespace DataConverter
{
    internal static class Console
    {
        public static TerminalControl Terminal { get; private set; }

        public static void SetTerminal(TerminalControl terminal)
        {
            Terminal = terminal;
        }

        public static void Print(string value, TerminalControl.LogType logType = TerminalControl.LogType.Info)
        {
            Terminal?.Append(value, logType);
        }
        public static void Print(string value, Brush color)
        {
            Terminal?.AppendRaw(value, color);
        }

        public static void PrintError(string value)
        {
            Print(value, TerminalControl.LogType.Error);
        }
        public static void PrintWarning(string value)
        {
            Print(value, TerminalControl.LogType.Warning);
        }
    }
}
