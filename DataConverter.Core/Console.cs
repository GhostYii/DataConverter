namespace DataConverter.Core
{
    public static class Console
    {
        private static Action<string> _defaultAction = msg => { };

        private static event Action<string> _print = _defaultAction;
        private static event Action<string> _warning = _defaultAction;
        private static event Action<string> _error = _defaultAction;

        public static void AddPrintListener(Action<string> action) { _print += action; }
        public static void RemovePrintListener(Action<string> action) { _print -= action; }

        public static void AddWarningListener(Action<string> action) { _warning += action; }
        public static void RemoveWarningListener(Action<string> action) { _warning -= action; }

        public static void AddErrorListener(Action<string> action) { _error += action; }
        public static void RemoveErrorListener(Action<string> action) { _error -= action; }

        public static void Clear()
        {
            foreach (var d in _print.GetInvocationList())
            {
                _print -= d as Action<string>;
            }
            foreach (var d in _warning.GetInvocationList())
            {
                _print -= d as Action<string>;
            }
            foreach (var d in _error.GetInvocationList())
            {
                _print -= d as Action<string>;
            }
        }

        internal static void Print(string message) { _print.Invoke(message); }
        internal static void PrintWarning(string message) { _warning.Invoke(message); }
        internal static void PrintError(string message) { _error.Invoke(message); }

    }
}
