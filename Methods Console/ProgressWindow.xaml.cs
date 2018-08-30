using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace Methods_Console
{
    /// <summary>
    /// Interaction logic for ProgressWindow.xaml
    /// </summary>
    public partial class ProgressWindow 
    {
        public string FileName { get; set; }
        public ProgressWindow(string filename)
        {
            InitializeComponent();
            FileName = filename;

        }
        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OpenFileWithWord(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer.exe", @"C:\BaaN-DAT\");
            System.Threading.Thread.Sleep(1000);
            Process.Start("winword.exe", FileName);
            Close();
        }

        private void OpenFileWithNotepad(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer.exe", @"C:\BaaN-DAT\");
            System.Threading.Thread.Sleep(1000);
            Process.Start("notepad.exe", FileName);
            Close();
        }
    }
}
