using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Outlook2010Sender.Views
{
    /// <summary>
    /// Interaction logic for SelectListPopup.xaml
    /// </summary>
    public partial class SelectListPopup : Window
    {
        public SelectListPopup()
        {
            InitializeComponent();
            
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.Multiselect = false;
            dlg.DefaultExt = ".msg";
            dlg.Filter = "Message Files (*.msg)|*.msg";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                FileLocation_tb.Text = filename;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            ThisAddIn.SendMails(FileLocation_tb.Text);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
