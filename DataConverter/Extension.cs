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

        public static T[] SubArray<T>(this T[] arr, int offset, int count)
        {
            if (count <= 0)
                return arr;

            return new ArraySegment<T>(arr, offset, count).ToArray();
        }
    }
}
