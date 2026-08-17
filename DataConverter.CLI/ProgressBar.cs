namespace DataConverter.CLI
{
    using Console = System.Console;

    internal sealed class ProgressBar : IDisposable
    {
        private readonly int _total;
        private readonly string _label;
        private readonly object _lock = new();

        private int _completed;
        private int _lastLineLength;
        private bool _isCompleted;

        public ProgressBar(int total, string label)
        {
            _total = Math.Max(0, total);
            _label = label;
            ConsoleOutput.AttachProgress(this);
            Draw(0);
        }

        public void Increment()
        {
            lock (_lock)
            {
                _completed = Math.Min(_total, _completed + 1);
                Draw(_completed);
            }
        }

        public void Complete()
        {
            lock (_lock)
            {
                if (_isCompleted)
                    return;

                if (_completed < _total)
                    Draw(_completed);
                ConsoleOutput.WriteProgress(Environment.NewLine);
                _lastLineLength = 0;
                _isCompleted = true;
                ConsoleOutput.DetachProgress();
            }
        }

        public void Dispose()
        {
            Complete();
        }

        private void Draw(int completed)
        {
            int barWidth = GetBarWidth();
            int filled = _total == 0 ? barWidth : (int)Math.Round(barWidth * completed / (double)_total);
            double percent = _total == 0 ? 100d : completed * 100d / _total;

            string line = $"\r{_label} [{new string('#', filled).PadRight(barWidth)}] " +
                          $"{percent,5:F1}% ({completed}/{_total})";

            ConsoleOutput.WriteProgress(line);

            int clear = _lastLineLength - line.Length;
            if (clear > 0)
                ConsoleOutput.WriteProgress(new string(' ', clear));

            _lastLineLength = line.Length;
        }

        private int GetBarWidth()
        {
            try
            {
                return Math.Clamp(Console.WindowWidth - 32, 10, 80);
            }
            catch
            {
                return 40;
            }
        }
    }
}
