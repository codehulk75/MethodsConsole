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
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;

namespace Methods_Console
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 


    public partial class MainWindow
    {


        int themenum = Properties.Settings.Default.WindowThemeNumber;
        int borderbrushnum = Properties.Settings.Default.StatusBarThemeNumber;
        Dictionary<int, string> themes = new Dictionary<int, string>();
        Dictionary<int, string> borderBrushes = new Dictionary<int, string>();
        List<Ci2Parser> ExportList = new List<Ci2Parser>();
        BeiBOM Bom;
        Assembly thisProgram;
        AssemblyName thisProgramName;
        Version thisProgramVersion;
        List<TextBox> lstProgTextBoxes = new List<TextBox>();
        List<ComboBox> lstProgComboBoxes = new List<ComboBox>();
        List<string> lstPasses = new List<string>(new string[] { "SMT 1", "SMT 2" });
        List<ComboBox> lstMachComboBoxes = new List<ComboBox>();
        List<Label> lstProgDateLabels = new List<Label>();
        Dictionary<int, string> dictMachines = new Dictionary<int, string>();
        public MainWindow()
        {
            thisProgram = Assembly.GetEntryAssembly();
            thisProgramName = thisProgram.GetName();
            thisProgramVersion = thisProgramName.Version;

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
            labelLoadingOne.Visibility = Visibility.Hidden;
            labelLoadingTwo.Visibility = Visibility.Hidden;
            tbLoadingInsOne.Visibility = Visibility.Hidden;
            tbLoadingInsTwo.Visibility = Visibility.Hidden;
            btnCreate.Visibility = Visibility.Hidden;
            dictMachines.Add(0, "FZ60XC");
            dictMachines.Add(1, "GC-60_1");
            dictMachines.Add(2, "GC-60_2");
            dictMachines.Add(3, "GI-14");
            dictMachines.Add(4, "GX-11");
            lstProgTextBoxes.Add(textBoxProg1);
            lstProgTextBoxes.Add(textBoxProg2);
            lstProgTextBoxes.Add(textBoxProg3);
            lstProgTextBoxes.Add(textBoxProg4);
            lstProgTextBoxes.Add(textBoxProg5);
            lstProgTextBoxes.Add(textBoxProg6);
            lstProgTextBoxes.Add(textBoxProg7);
            lstProgTextBoxes.Add(textBoxProg8);
            lstProgComboBoxes.Add(comboBoxProg1);
            lstProgComboBoxes.Add(comboBoxProg2);
            lstProgComboBoxes.Add(comboBoxProg3);
            lstProgComboBoxes.Add(comboBoxProg4);
            lstProgComboBoxes.Add(comboBoxProg5);
            lstProgComboBoxes.Add(comboBoxProg6);
            lstProgComboBoxes.Add(comboBoxProg7);
            lstProgComboBoxes.Add(comboBoxProg8);
            lstMachComboBoxes.Add(comboBoxMach1);
            lstMachComboBoxes.Add(comboBoxMach2);
            lstMachComboBoxes.Add(comboBoxMach3);
            lstMachComboBoxes.Add(comboBoxMach4);
            lstMachComboBoxes.Add(comboBoxMach5);
            lstMachComboBoxes.Add(comboBoxMach6);
            lstMachComboBoxes.Add(comboBoxMach7);
            lstMachComboBoxes.Add(comboBoxMach8);
            lstProgDateLabels.Add(labelProg1);
            lstProgDateLabels.Add(labelProg2);
            lstProgDateLabels.Add(labelProg3);
            lstProgDateLabels.Add(labelProg4);
            lstProgDateLabels.Add(labelProg5);
            lstProgDateLabels.Add(labelProg6);
            lstProgDateLabels.Add(labelProg7);
            lstProgDateLabels.Add(labelProg8);

            foreach (TextBox box in lstProgTextBoxes)
            {
                box.Visibility = Visibility.Hidden;
                box.IsReadOnly = true;
            }
            foreach(ComboBox cb in lstProgComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
                cb.Items.Add(lstPasses[0]);
                cb.Items.Add(lstPasses[1]);
                cb.SelectedIndex = 0;
            }
            foreach(ComboBox cb in lstMachComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
                cb.Items.Add(dictMachines[0]);
                cb.Items.Add(dictMachines[1]);
                cb.Items.Add(dictMachines[2]);
                cb.Items.Add(dictMachines[3]);
                cb.Items.Add(dictMachines[4]);
                cb.SelectedIndex = 0;
            }
            foreach (Label lb in lstProgDateLabels)
            {
                lb.Visibility = Visibility.Hidden;
            }

        }

        private void ClearData()
        {
            ExportList.Clear();
            Bom = null;
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

        private void PrepareSetupSheet()
        {          
            Tuple<bool, List<string>, List<string>> tDetectInfo;
            tDetectInfo = AutoDetectExportInfo();
            ExportList = ExportList.OrderBy(p => p.Pass).ThenBy(m => m.MachineName).ToList();
            PopulateUIBoxes();
            if (tDetectInfo.Item1 == false)
            {
                MessageBox.Show("Add function for failed detection.", "HURRY UP!! No pressure...");
            }

        }

        private void PopulateUIBoxes()
        {
            int i = 0;
            btnCreate.Visibility = Visibility.Visible;
            labelLoadingOne.Visibility = Visibility.Visible;
            labelLoadingTwo.Visibility = Visibility.Visible;
            tbLoadingInsOne.Visibility = Visibility.Visible;
            tbLoadingInsTwo.Visibility = Visibility.Visible;
            foreach( Ci2Parser program in ExportList)
            {
                lstProgTextBoxes[i].Text = program.ProgramName;
                lstProgTextBoxes[i].Visibility = Visibility.Visible;
                lstProgComboBoxes[i].Visibility = Visibility.Visible;
                lstMachComboBoxes[i].Visibility = Visibility.Visible;
                lstProgDateLabels[i].Content = program.DateCreated;
                lstProgDateLabels[i].Visibility = Visibility.Visible;              
                if (program.Pass.Equals(lstPasses[0]))
                { 
                    lstProgComboBoxes[i].SelectedIndex = 0;
                }
                else if (program.Pass.Equals(lstPasses[1]))
                {
                    lstProgComboBoxes[i].SelectedIndex = 1;
                }
                foreach(var item in dictMachines)
                {
                    if (program.MachineName.Equals(item.Value))
                        lstMachComboBoxes[i].SelectedIndex = item.Key;
                }
                ++i;
            }

        }
        private void CreateBaanSetupSheet()
        {
            int iCurrentPageNum = 1;
            string strSsFilePath = "C:\\BaaN-DAT\\" + Bom.AssemblyName + "_" + Bom.Rev + ".rtf";
            CreateRtfDoc(strSsFilePath);
            iCurrentPageNum = WriteProgramData(strSsFilePath);



            RtfDocWriteLastLine(strSsFilePath, iCurrentPageNum);
        }

        private void CreateRtfDoc(string strFilePath)
        {
            string strRtfFirstLine = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Courier New;}}\viewkind4\uc1\pard\tx2160\tx3600\tx4320\tx7200\margl360\margr360\margt360\margb360 {\f0\fs20";
            string strInitialHeader = @"\tab Assembly:\tab " + Bom.AssemblyName + "\n"
                       + @"\par \tab BOM Rev:\tab " + Bom.Rev + "\n"
                       + @"\par \tab Listing Date:\tab " + Bom.DateOfListing + "\n"
                       + @"\par _______________________________________________________________________________________________" + "\n\n";                           
            using (StreamWriter writer = new StreamWriter(strFilePath, false))
            {
                writer.WriteLine(strRtfFirstLine);
                writer.WriteLine(strInitialHeader);
            }
            
        }

        private int WriteProgramData(string strFilePath)
        {
            int iCurrentPageNumber = 1;

            foreach (Ci2Parser export in ExportList)
            {
                string strInstructions = tbLoadingInsOne.Text;
                if (export.Pass.Equals(lstPasses[1]))
                    strInstructions = tbLoadingInsTwo.Text;
                WriteNewProgramHeader(strFilePath, export, strInstructions);
            }

            using (StreamWriter writer = new StreamWriter(strFilePath, true))
            {
                writer.WriteLine("BLAHHHH");
            }
            return iCurrentPageNumber;
        }

        private void WriteNewProgramHeader(string strFilePath, Ci2Parser ci2, string strLoadingInstructions)
        {
            string strNewProgramHeader = @"\par Program: " + ci2.ProgramName + @"\tab Date: " + ci2.DateCreated + @"\tab \tab " + ci2.MachineName + @"\tab Side: " + ci2.Pass + "\n"
                                        + @"\par Loading Instructions: " + strLoadingInstructions + "\n"
                                        + "\\par \n\\par Part Number\\tab Description\n"
                                        + @"\par\tab Feeder\tab Qty\tab Reference Designators" + "\n"
                                        + @"\par _______________________________________________________________________________________________" + "\n";
            using (StreamWriter writer = new StreamWriter(strFilePath, true))
            {
                writer.WriteLine(strNewProgramHeader);
                writer.WriteLine("\\par \\tab some progrm info....\n");
            }

        }

        private int WriteMidProgramFooterHeader(string strFilePath, int iCurrentPageNum)
        {
            int i = iCurrentPageNum;
            string strFooterHeader = "\n\\par\n"
                                    + @"\par " + DateTime.Now.ToShortDateString() + @" \tab Page  " + iCurrentPageNum.ToString() + @" \tab Version " + thisProgramVersion.ToString() + @"\page \tab Assembly:\tab " + Bom.AssemblyName + "\n"
                                    + @"\par \tab BOM Rev:\tab " + Bom.Rev + "\n"
                                    + @"\par \tab Listing Date:\tab " + Bom.DateOfListing + "\n"
                                    + "\\par _______________________________________________________________________________________________\n\n";
            using (StreamWriter writer = new StreamWriter(strFilePath, true))
            {
                writer.WriteLine(strFooterHeader);
            }
            return ++i;
        }
        private void RtfDocWriteLastLine(string strFilePath, int iLastPageNum)
        {
            string rtfLastLine = @"\par " + DateTime.Now.ToShortDateString() + @" \tab Page  " + iLastPageNum.ToString() + " \tab Version " + thisProgramVersion.ToString() + @"\page }}";
            using (StreamWriter writer = new StreamWriter(strFilePath, true))
            {
                writer.WriteLine(rtfLastLine);
            }
        }
        private void CreateAgileSetupSheet()
        {

        }
        private Tuple<bool, List<string>, List<string>> AutoDetectExportInfo()
        {
            //
            //Try to auto detect Pass and machine name based on ci2 file name
            //
            bool bDectectionSuccessful = false;
            List<string> failedMachineDetection = new List<string>();
            List<string> failedPassDetection = new List<string>();
            Regex reBot = new Regex(@"\bBOT\b", RegexOptions.IgnoreCase);
            Regex reTop = new Regex(@"\bTOP\b", RegexOptions.IgnoreCase);
            Regex reFuzion = new Regex(@"\s+FZ60XC\s+",RegexOptions.IgnoreCase);
            Regex reGCOne = new Regex(@"\s+(GC-?60-?_?1?)\s+", RegexOptions.IgnoreCase);
            Regex reGCTwo = new Regex(@"\s+(GC-?60-?_2?)\s+", RegexOptions.IgnoreCase);
            Regex reGI = new Regex(@"\s+(GI-?14-?_?\d?)\s+", RegexOptions.IgnoreCase);
            Regex reGX = new Regex(@"\s+(GX-?11)\s+", RegexOptions.IgnoreCase);

            foreach (Ci2Parser parser in ExportList)
            {
                string filename = parser.FileName;
                if (reBot.IsMatch(filename))
                {
                    parser.Pass = lstPasses[0];
                }
                else if (reTop.IsMatch(filename))
                {
                    parser.Pass = lstPasses[1];
                }
                else
                {
                    failedPassDetection.Add(parser.FileName);
                }
                if (reFuzion.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[0];
                }
                else if (reGCOne.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[1];
                }
                else if (reGCTwo.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[2];
                }
                else if (reGI.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[3];
                }
                else if (reGX.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[4];
                }
                else
                {
                    failedMachineDetection.Add(parser.FileName);
                }
            }
            if (failedMachineDetection.Count == 0 && failedPassDetection.Count == 0)
                bDectectionSuccessful = true;

            return Tuple.Create(bDectectionSuccessful, failedPassDetection, failedMachineDetection);

        }

        private void setupsheetbutton_Click(object sender, RoutedEventArgs e)
        {
            ClearData();
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
                        ExportList.Add(parser);
                    }
                }             
            }
            LoadBOM();
            PrepareSetupSheet();
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            if (Bom.HasRouting)
                CreateBaanSetupSheet();
            else
                CreateAgileSetupSheet();
        }

        private void comboBoxMach1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[0].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[1].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[2].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[3].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach5_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[4].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach6_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[5].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach7_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[6].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach8_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[7].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[0].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[1].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[2].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[3].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg5_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[4].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg6_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[5].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg7_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[6].Pass = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg8_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible)
                ExportList[7].Pass = (sender as ComboBox).SelectedItem.ToString();
        }
    }
}
