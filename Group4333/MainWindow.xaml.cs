using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            _4333_MonichArtem info = new _4333_MonichArtem();
            info.Show();
        }
    }
}