using System;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace DataConverter
{
    /// <summary>
    /// Interaction logic for TerminalControl.xaml
    /// </summary>
    public partial class TerminalControl : UserControl
    {
        public enum LogType
        {
            None,
            Info,
            Warning,
            Error
        }

        public bool AutoScrollToEnd { get; set; }

        private FlowDocument _terminalDoc = null;

        public TerminalControl()
        {
            InitializeComponent();
            
            _terminalDoc = new FlowDocument();
            outputBox.Document = _terminalDoc;
        }

        public void Append(string text, LogType type = LogType.Info)
        {
            Paragraph paragraph = new Paragraph();
            Run timestamp = new Run(string.Format("[{0}]", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")));
            timestamp.Foreground = Brushes.White;
            paragraph.Inlines.Add(timestamp);

            if (type != LogType.None)
            {
                Run typePrefix = new Run(string.Format("[{0}]", type));
                typePrefix.Foreground = type.ToBrush();
                paragraph.Inlines.Add(typePrefix);
            }

            Run content = new Run($": {text}");
            content.Foreground = type.ToBrush();
            paragraph.Inlines.Add(content);
            
            _terminalDoc.Blocks.Add(paragraph);

            if (AutoScrollToEnd)
                outputBox.ScrollToEnd();
        }

        public void AppendRaw(string text, Brush brush = null)
        {
            Paragraph paragraph = new Paragraph();
            Run content = new Run(text);
            content.Foreground = brush ?? Brushes.White;
            paragraph.Inlines.Add(content);

            _terminalDoc.Blocks.Add(paragraph);
            if (AutoScrollToEnd)
                outputBox.ScrollToEnd();
        }

        public void Clear()
        {
            _terminalDoc.Blocks.Clear();
        }
    }
}
