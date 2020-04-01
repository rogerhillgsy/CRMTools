using System;
using System.Windows;
using System.Windows.Controls;

namespace Microsoft.Crm.Sdk.Samples
{
    /// <summary>
    /// Interaction logic for ErrorWindow.xaml
    /// </summary>
    public partial class ErrorWindow : Window
    {
        public String Message { get; set; }

        public ErrorWindow()
        {
            InitializeComponent();
        }

        public ErrorWindow(string errorMessage)
        {
            InitializeComponent();

            Message = errorMessage;

            this.DataContext = Message;
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
