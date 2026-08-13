namespace DataConverter.CLI
{
    using Console = System.Console;

    internal static class ConsoleOutput
    {
        private static readonly object Lock = new();
        private static readonly List<(string Text, ConsoleColor? Color)> Pending = new();
        private static ProgressBar? _progress;

        public static void AttachProgress(ProgressBar progress)
        {
            lock (Lock)
            {
                _progress = progress;
                Pending.Clear();
            }
        }

        public static void WriteProgress(string text)
        {
            lock (Lock)
            {
                Console.Write(text);
            }
        }

        public static void DetachProgress()
        {
            lock (Lock)
            {
                _progress = null;
                foreach (var (text, color) in Pending)
                    WriteLineNow(text, color);
                Pending.Clear();
            }
        }

        public static void WriteLine(string message, ConsoleColor? color = null)
        {
            lock (Lock)
            {
                if (_progress != null)
                {
                    Pending.Add((message, color));
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
