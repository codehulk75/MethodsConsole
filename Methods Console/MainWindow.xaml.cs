using ModernChrome;
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

namespace Methods_Console
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        int themenum = 1;
        public MainWindow()
        {
            InitializeComponent();
            ThemeManager.ThemeChanged += (sender, args) =>
            {
                if (args.Theme != "Light")
                {
                    TasksText.Foreground = Brushes.White;
                    setupsheetbutton.Foreground = Brushes.White;
                    famcheckbutton.Foreground = Brushes.White;
                    themebutton.Foreground = Brushes.White;
                    comparebutton.Foreground = Brushes.White;
                }
                else
                {
                    TasksText.Foreground = Brushes.Black;
                    setupsheetbutton.Foreground = Brushes.Black;
                    famcheckbutton.Foreground = Brushes.Black;
                    themebutton.Foreground = Brushes.Black;
                    comparebutton.Foreground = Brushes.Black;
                }
            };
        }

        private void button_Copy2_Click(object sender, RoutedEventArgs e)
        {
            ++themenum;
            switch (themenum)
            {
                case 1:
                    ThemeManager.ChangeTheme(Application.Current, "LightBlue");
                    break;
                case 2:
                    ThemeManager.ChangeTheme(Application.Current, "DarkBlue");
                    break;
                case 3:
                    ThemeManager.ChangeTheme(Application.Current, "Light");
                    break;
                case 4:
                    ThemeManager.ChangeTheme(Application.Current, "Dark");
                    break;
                case 5:
                    ThemeManager.ChangeTheme(Application.Current, "Blend");
                    break;
                default:
                    themenum = 1;
                    ThemeManager.ChangeTheme(Application.Current, "LightBlue");
                    break;
            }
        }
    }
}
