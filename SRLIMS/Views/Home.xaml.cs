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
using System.Windows.Shapes;

namespace SRLIMS.Views
{
    /// <summary>
    /// Lógica de interacción para Window1.xaml
    /// </summary>
    public partial class Home : UserControl
    {
        public Home()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OpenSelectDataSourceFrame(object sender, RoutedEventArgs e)
        {

            try
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (Application.Current.MainWindow is MainWindow mainWindow && mainWindow.MainFrame != null)
                    {

                        var newView = new SelectDataSource();

                        if (mainWindow.MainFrame.Content is IDisposable oldView)
                        {
                            oldView.Dispose();
                        }

                        mainWindow.MainFrame.Content = newView;
                    }
                }), System.Windows.Threading.DispatcherPriority.Normal);
            }
            catch (Exception ex) {

                MessageBox.Show($"Error when change the window{ex.Message}");

            }



            var mainWindow = Application.Current.MainWindow as MainWindow;

            mainWindow.MainFrame.Content = new SelectDataSource();  
            

        }
    }
}
