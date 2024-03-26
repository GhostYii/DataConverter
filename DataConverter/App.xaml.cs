using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

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

            TypeParser.LoadParser();
        }
    }
}
