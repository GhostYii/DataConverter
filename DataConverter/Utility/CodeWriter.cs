using System.Linq;
using System.Text;

namespace DataConverter.Utility
{
    internal class CodeWriter
    {
        public enum TargetType
        {
            Class,
            Struct,
        }

        private StringBuilder _buffer = null;
        private int _indentLevel = 0;
        private uint _spacePerIndent = 4;

        public CodeWriter(uint spacePerIndent = 4, TargetType type = TargetType.Class)
        {
            _buffer = new StringBuilder();
            _indentLevel = 0;
            _spacePerIndent = spacePerIndent;
        }

        public void Write(string content)
        {
            _buffer.Append(content);
        }

        public void WriteFormat(string format, params object?[] args)
        {
            _buffer.AppendFormat(format, args);
        }

        public void WriteLine(string text = "")
        {
            if (!text.All(char.IsWhiteSpace))
            {
                WriteIndent();
                _buffer.Append(text);
            }
            _buffer.AppendLine();
        }

        public void BeginBlock()
        {
            WriteIndent();
            _buffer.AppendFormat("{{{0}", System.Environment.NewLine);
            ++_indentLevel;
        }

        public void EndBlock()
        {
            --_indentLevel;
            WriteIndent();
            _buffer.AppendFormat("}}{0}", System.Environment.NewLine);
        }

        public void WriteIndent()
        {
            for (var i = 0; i < _indentLevel; ++i)
            {
                for (var n = 0; n < _spacePerIndent; ++n)
                    _buffer.Append(' ');
            }
        }

        public override string ToString()
        {
            return _buffer.ToString();
        }

    }
}
