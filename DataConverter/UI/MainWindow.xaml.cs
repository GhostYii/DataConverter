using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DataConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<int, int> _tabHeightDict = new Dictionary<int, int>()
        {
            [0] = 130,
            [1] = 300
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnTabChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_tabHeightDict.ContainsKey(tabControl.SelectedIndex))
                return;
        }

        private void OnMainWindowClosed(object sender, EventArgs e)
        {
            App app = (App.Current as App);
            app.ConsoleWindow.Close();
        }
    }
}
