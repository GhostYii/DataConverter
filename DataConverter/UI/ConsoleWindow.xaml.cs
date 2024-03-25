using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace DataConverter
{
    /// <summary>
    /// Interaction logic for ConsoleWindow.xaml
    /// </summary>
    public partial class ConsoleWindow : Window
    {
        private int _maxHistoryCount = 100;
        private int _curHistoryIndex = 0;
        private List<string> _historyList = new List<string>();

        public ConsoleWindow()
        {
            InitializeComponent();

            terminal.AutoScrollToEnd = true;

            Commands.Terminal = terminal;
            Commands.InitialCommandByType(typeof(CustomCommands));

            Console.SetTerminal(terminal);
        }

        private void OnCloseButtonClicked(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OnMinButtonClicked(object sender, RoutedEventArgs e)
        {

        }

        private void OnMaxButtonClicked(object sender, RoutedEventArgs e)
        {

        }

        private void DragWithTitleBar(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void OnCmdBoxKeyDown(object sender, KeyEventArgs e)
        {            
            if (e.Key != Key.Enter)
                return;

            if (string.IsNullOrEmpty(cmdBox.Text))
            {
                terminal.Append(string.Empty, TerminalControl.LogType.None);
                return;
            }

            List<string> cmd = new List<string>(cmdBox.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries));
            if (cmd.Count == 0)
                return;

            terminal.Append(cmdBox.Text);            
            Commands.Execute(cmd[0], cmd.GetRange(1, cmd.Count - 1).ToArray());

            if (_historyList.Count >= _maxHistoryCount)
                _historyList.RemoveAt(0);

            _historyList.Add(cmdBox.Text);

            _curHistoryIndex = _historyList.Count;
            cmdBox.Clear();
        }

        private void OnCmdBoxPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (_historyList.Count == 0)
                return;

            if (e.Key == Key.Up)
            {
                cmdBox.Clear();
                _curHistoryIndex = Math.Clamp(--_curHistoryIndex, 0, _historyList.Count - 1);
                cmdBox.Text = _historyList[_curHistoryIndex];
                cmdBox.Select(cmdBox.Text.Length, 0);
            }
            else if (e.Key == Key.Down)
            {
                cmdBox.Clear();
                _curHistoryIndex = Math.Clamp(++_curHistoryIndex, 0, _historyList.Count - 1);
                cmdBox.Text = _historyList[_curHistoryIndex];
                cmdBox.Select(cmdBox.Text.Length, 0);
            }            
        }

        private void OnCmdBoxTextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            
        }
    }
}
