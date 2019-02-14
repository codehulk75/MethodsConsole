
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
        BeiBOM BomOne;
        BeiBOM BomTwo;
        public BomCompareWindow()
        {
            BomOne = null;
            BomTwo = null;
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
            if(BomOne != null && BomTwo != null && BomOne.IsValid && BomTwo.IsValid)
            {
                BOMComparer bomComparer = new BOMComparer(BomOne, BomTwo);
                bool BomsMatch = bomComparer.CompareBomData();

            }
            else
            {
                MessageBox.Show("You need 2 valid BOMs to perform a BOM Compare.", "Insufficient Data", MessageBoxButton.OK, MessageBoxImage.Hand);
            }

        }

        private void BomOneButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "BOM File|*.xls;*.csv;*.txt|All Files|*.*";
            dlg.Title = "Load BOM File #1";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                if (BomOne != null)
                    BomOne = null;
                BomOne = new BeiBOM(dlg.FileName);
                if (BomOne.IsValid == false)
                {
                    MessageBox.Show("Invalid BOM");
                    tbBom1FileName.Text = "File: ";
                    tbBom1FileName.ToolTip = "";
                    tblkAssemblyOne.Text = "Assembly: ";
                    tblkAssemblyOne.ToolTip = "";
                    tblkBom1Rev.Text = "Rev: ";
                    tblkBom1Date.Text = "Date: ";
                    BomOne = null;
                    return;
                }
                tbBom1FileName.Text = "File:  " + BomOne.FileName;
                tbBom1FileName.ToolTip = BomOne.FullFilePath;
                tblkAssemblyOne.Text = "Assembly:  " + BomOne.AssemblyName;
                tblkAssemblyOne.ToolTip = BomOne.AssyDescription;
                tblkBom1Rev.Text = "Rev:  " + BomOne.Rev;
                tblkBom1Date.Text = "Date:  " + BomOne.DateOfListing;

            }
        }
        private void BomTwoButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "BOM File|*.xls;*.csv;*.txt|All Files|*.*";
            dlg.Title = "Load BOM File #2";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                if (BomTwo != null)
                    BomTwo = null;
                BomTwo = new BeiBOM(dlg.FileName);
                if (BomTwo.IsValid == false)
                {
                    MessageBox.Show("Invalid BOM");
                    tbBom2FileName.Text = "File: ";
                    tbBom2FileName.ToolTip = "";
                    tblkAssemblyTwo.Text = "Assembly: ";
                    tblkAssemblyTwo.ToolTip = "";
                    tblkBom2Rev.Text = "Rev: ";
                    tblkBom2Date.Text = "Date: ";
                    BomTwo = null;
                    return;
                }
                tbBom2FileName.Text = "File:  " + BomTwo.FileName;
                tbBom2FileName.ToolTip = BomTwo.FullFilePath;
                tblkAssemblyTwo.Text = "Assembly:  " + BomTwo.AssemblyName;
                tblkAssemblyTwo.ToolTip = BomTwo.AssyDescription;
                tblkBom2Rev.Text = "Rev:  " + BomTwo.Rev;
                tblkBom2Date.Text = "Date:  " + BomTwo.DateOfListing;

            }
        }
    }
}
