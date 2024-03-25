using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace DataConverter
{
    internal static class Extension
    {

        public static SolidColorBrush ToBrush(this TerminalControl.LogType type)
        {
            switch (type)
            {
                case TerminalControl.LogType.Info:
                    return Brushes.White;
                case TerminalControl.LogType.Warning:
                    return Brushes.Yellow;
                case TerminalControl.LogType.Error:
                    return Brushes.Red;
                default:
                    return Brushes.White;
            }
        }

        // ref type return object
        public static Type ToType(this CellType cellType)
        {
            switch (cellType.type)
            {
                case ValueType.Int:
                    return typeof(int);
                case ValueType.Float:
                    return typeof(float);
                case ValueType.String:
                    return typeof(string);
                case ValueType.Object:
                    return Type.GetType(cellType.objName);
                case ValueType.Ref:
                    return typeof(object);
            }

            return null;
        }
    }
}
