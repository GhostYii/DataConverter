using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using DataConverter.Core;

namespace DataConverter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public ConsoleWindow ConsoleWindow { get; private set; }

        private void OnApplicationStartup(object sender, StartupEventArgs e)
        {
            ConsoleWindow = new ConsoleWindow();
            ConsoleWindow.Show();

            DataConverter.Core.Console.AddPrintListener(Console.Print);
            DataConverter.Core.Console.AddWarningListener(Console.PrintWarning);
            DataConverter.Core.Console.AddErrorListener(Console.PrintError);

            DataConverter.Core.DC.RegisterTypes(typeof(TestObject));
        }
    }
}
