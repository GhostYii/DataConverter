namespace DataConverter.CLI
{
    using Console = System.Console;

    internal static class ConsoleOutput
    {
        private static readonly object _lock = new();
        private static readonly List<(string Text, ConsoleColor? Color)> _pending = new();
        private static ProgressBar? _progress;

        public static void AttachProgress(ProgressBar progress)
        {
            lock (_lock)
            {
                _progress = progress;
                _pending.Clear();
            }
        }

        public static void WriteProgress(string text)
        {
            lock (_lock)
            {
                Console.Write(text);
            }
        }

        public static void DetachProgress()
        {
            lock (_lock)
            {
                _progress = null;
                foreach (var (text, color) in _pending)
                    WriteLineNow(text, color);
                _pending.Clear();
            }
        }

        public static void WriteLine(string message, ConsoleColor? color = null)
        {
            lock (_lock)
            {
                if (_progress != null)
                {
                    _pending.Add((message, color));
                    return;
                }

                WriteLineNow(message, color);
            }
        }

        private static void WriteLineNow(string message, ConsoleColor? color)
        {
            if (color.HasValue)
                Console.ForegroundColor = color.Value;

            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
