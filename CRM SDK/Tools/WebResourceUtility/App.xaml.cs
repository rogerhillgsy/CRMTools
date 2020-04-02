using System;
using System.Collections.Generic;
using System.Windows;



namespace Microsoft.Crm.Sdk.Samples
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {       
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            MainWindow mainWindow = new MainWindow();
            MainWindowViewModel viewModel = new MainWindowViewModel();            

            mainWindow.DataContext = viewModel;
            mainWindow.Show();            
        }

        public App()
        {            
        }
    }
}
