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
        int themenum = Properties.Settings.Default.WindowThemeNumber;
        int borderbrushnum = Properties.Settings.Default.StatusBarThemeNumber;
        Dictionary<int, string> themes = new Dictionary<int, string>();
        Dictionary<int, string> borderBrushes = new Dictionary<int, string>();
        List<FileParser> ci2List = new List<FileParser>();
        BeiBOM Bom;
        public MainWindow()
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
            ThemeManager.ThemeChanged += (sender, args) =>
            {
                SetTheme(args.Theme, borderbrushnum);
            };
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
        }

        private void SetTheme(String strThemeColor, int nBorderColor)
        {
            BorderBrush = Application.Current.FindResource($"StatusBar{borderBrushes[nBorderColor]}BrushKey") as SolidColorBrush;
            SetBrushes(strThemeColor);
            int colorKey = themes.FirstOrDefault(x => x.Value == strThemeColor).Key;
            Properties.Settings.Default.StatusBarThemeNumber = nBorderColor;
            Properties.Settings.Default.WindowThemeNumber = colorKey;
            Properties.Settings.Default.Save();
        }

        private void button_theme_Click(object sender, RoutedEventArgs e)
        {
            ++themenum;
            if (themenum > 5)
                themenum = 1;
            ThemeManager.ChangeTheme(Application.Current, themes[themenum]);
        }

        private void themebutton_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ++borderbrushnum;
            if (borderbrushnum > 4)
                borderbrushnum = 1;
            SetTheme(themes[themenum], borderbrushnum);
        }

        private void LoadBOM()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "BOM Explosion/BAAN BOM|*.xls;*xlsx;*.csv;*.txt";
            dlg.Multiselect = false;
            dlg.Title = "Load the BOM File";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                Bom = new BeiBOM(dlg.FileName);
                if (Bom.IsValid)
                    MessageBox.Show(Bom.FileName + " is a valid BOM.");
                else
                    MessageBox.Show("Invalid BOM");
            }
        }

        private void CreateSetupSheet()
        {
            MessageBox.Show("Workin on CreateSetupSheet()", "not dun");


        }

        private void setupsheetbutton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Multiselect = true;
            dlg.Title = "Load the Program Exports";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string[] files = dlg.FileNames;
                foreach (var filename in files)
                {
                    FileParserFactory factory = new Ci2ParserFactory(filename);
                    Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                    if(parser != null)
                    {
                        ci2List.Add(parser);
                    }
                }
            }
            LoadBOM();
            CreateSetupSheet();
        }
    }
}
