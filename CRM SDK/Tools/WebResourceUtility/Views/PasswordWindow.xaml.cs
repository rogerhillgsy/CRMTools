using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Microsoft.Crm.Sdk.Samples
{
    /// <summary>
    /// Interaction logic for PasswordWindow.xaml
    /// </summary>
    public partial class PasswordWindow : Window
    {       
        public PasswordWindow()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(this.PasswordText.Password))
            {
                this.DialogResult = true;
            }
            else
            {
                this.DialogResult = false;
            }
        }

        public string GetPassword()
        {
            return this.PasswordText.Password; 
        }
    }
}
