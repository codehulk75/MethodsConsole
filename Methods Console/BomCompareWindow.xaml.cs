
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
using System.Windows.Shapes;

namespace Methods_Console
{
    /// <summary>
    /// Interaction logic for BomCompareWindow.xaml
    /// </summary>
    public partial class BomCompareWindow
    {
        int themenum = Properties.Settings.Default.WindowThemeNumber;
        int borderbrushnum = Properties.Settings.Default.StatusBarThemeNumber;
        Dictionary<int, string> themes = new Dictionary<int, string>();
        Dictionary<int, string> borderBrushes = new Dictionary<int, string>();
        public BomCompareWindow()
        {
            InitializeComponent();
            InitializeThemeColors();
            try
            {
                ThemeManager.ChangeTheme(Application.Current, themes[themenum]);
                BorderBrush = Application.Current.FindResource($"StatusBar{borderBrushes[borderbrushnum]}BrushKey") as SolidColorBrush;
                SetBrushes(themes[themenum]);
            }
            catch (Exception)
            {
                MessageBox.Show("Error setting window skin.\nMainWindow()->ThemeManager.ChangeTheme())");
            }
        }
        private void InitializeThemeColors()
        {
            themes.Add(1, "LightBlue");
            themes.Add(2, "DarkBlue");
            themes.Add(3, "Light");
            themes.Add(4, "Dark");
            themes.Add(5, "Blend");
            borderBrushes.Add(1, "Blue");
            borderBrushes.Add(2, "Orange");
            borderBrushes.Add(3, "Purple");
            borderBrushes.Add(4, "Green");
        }
        private void SetBrushes(String strThemeColor)
        {
            if (strThemeColor != "Light")
            {
                StripPrefixesCheckbox.Foreground = Brushes.White;
                tbBom1FileName.Foreground = Brushes.White;
                tbBom2FileName.Foreground = Brushes.White;
                tblkAssemblyOne.Foreground = Brushes.White;
                tblkAssemblyTwo.Foreground = Brushes.White;
                tblkBom1Rev.Foreground = Brushes.White;
                tblkBom2Rev.Foreground = Brushes.White;
                tblkBom1Date.Foreground = Brushes.White;
                tblkBom2Date.Foreground = Brushes.White;
            }
            else
            {
                StripPrefixesCheckbox.Foreground = Brushes.Black;
                tbBom1FileName.Foreground = Brushes.Black;
                tbBom2FileName.Foreground = Brushes.Black;
                tblkAssemblyOne.Foreground = Brushes.Black;
                tblkAssemblyTwo.Foreground = Brushes.Black;
                tblkBom1Rev.Foreground = Brushes.Black;
                tblkBom2Rev.Foreground = Brushes.Black;
                tblkBom1Date.Foreground = Brushes.Black;
                tblkBom2Date.Foreground = Brushes.Black;
            }
        }

        private void AnalyzeButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BomOneButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BomTwoButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
